require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'json'
require 'stringio'
require 'active_support/time'

APPLICATION_NAME = 'Google Sheets API Ruby Integration'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

# Files to store assignment counts and assigned events
@assignment_counts_file = 'assignment_counts.csv'
@assigned_events_file = 'assigned_events.csv'

# Initialize global data as nested hashes keyed by team
@assignment_counts = Hash.new { |hash, key| hash[key] = Hash.new(0) }
@assigned_events = Hash.new { |hash, key| hash[key] = {} }

# Load assignment counts and assigned events from files if they exist
load_assignment_counts
load_assigned_events

# Google Sheets setup
def setup_google_sheets
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize
  service
end

# Google Sheets authorization using a service account
def authorize
  credentials = JSON.parse(ENV['GOOGLE_SHEETS_CREDENTIALS'])
  Google::Auth::ServiceAccountCredentials.make_creds(
    json_key_io: StringIO.new(credentials.to_json),
    scope: SCOPE
  )
end

# Method to clear team data in Google Sheets before updating
def clear_google_sheet_data(service, spreadsheet_id, range)
  clear_request = Google::Apis::SheetsV4::ClearValuesRequest.new
  service.clear_values(spreadsheet_id, range, clear_request)
end

# Method to fetch existing data from Google Sheets
def fetch_existing_data(service, spreadsheet_id, range)
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  response.values || []
end

# Method to sort the Google Sheet by date (assuming date is in the third column)
def sort_google_sheet_by_date(service, spreadsheet_id, sheet_id)
  sort_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
    requests: [
      {
        sort_range: {
          range: {
            sheet_id:,
            start_row_index: 1, # Skip header row
            start_column_index: 0,
            end_column_index: 6 # Assuming data goes up to column F
          },
          sort_specs: [
            {
              dimension_index: 2, # Date column index (third column)
              sort_order: 'ASCENDING'
            }
          ]
        }
      }
    ]
  )
  service.batch_update_spreadsheet(spreadsheet_id, sort_request)
end

# Method to write team data to Google Sheets
def write_team_data_to_individual_sheets(service, team, data)
  headers = ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
  values = [headers] + data
  range = 'Sheet1!A1:F'
  value_range = Google::Apis::SheetsV4::ValueRange.new(values:)
  service.update_spreadsheet_value(team[:spreadsheet_id], range, value_range, value_input_option: 'RAW')
end

# Define an exclusion list for events that do not require a locker room monitor
EXCLUSION_LIST = [
  'Skills Off Ice', 'Dryland', 'Goalie Training', 'Off Ice', 'Conditioning', 'Meeting', 'Goalie', 'LRM', 'Tournament', 'Pictures'
]

# Define teams and configurations
TEAMS = [
  {
    name: '12A',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019',
    family_names: %w[Becker Hastings Opel Gorgos Larsen Anderson Campos Powell Tousignant Marshall Johnson Wulff Orstad
                     Mulcahey],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_12A']
  },
  {
    name: '12B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603021',
    family_names: %w[Baer Bimberg Chanthavongsa Hammerstrom Kremer Lane Oas Perpich Ray Reinke Silva-Hammer],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_12B1']
  },
  {
    name: '10B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603022',
    family_names: %w[Baer Bowman Hopper Houghtaling Johnson Larsen Markfort Marshall Nanninga Orstad Willey Williamson],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_10B1']
  },
  {
    name: '10B2',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603023',
    family_names: %w[Curry Engholm Froberg Harpel Johnson Mckinnon Oprenchak M-Reberg B-Reberg Sauer Smith Woods],
    spreadsheet_id: ENV['GOOGLE_SHEET_ID_10B2']
  }
]

# Load assignment counts from file
def load_assignment_counts
  return unless File.exist?(@assignment_counts_file)

    CSV.foreach(@assignment_counts_file, headers: true) do |row|
      team_name = row['Team']
      family = row['Family']
      count = row['Count'].to_i
      @assignment_counts[team_name][family] = count
    end
end

# Load assigned events from file
def load_assigned_events
  return unless File.exist?(@assigned_events_file)

    CSV.foreach(@assigned_events_file, headers: true) do |row|
      team_name = row['Team']
      event_id = row['EventID']
      monitor = row['Locker Room Monitor']
      @assigned_events[team_name][event_id] = monitor
    end
end

# Create an event for the iCal feed
def create_ical_event(start_time, end_time, summary, description)
  event = Icalendar::Event.new
  event.dtstart = Icalendar::Values::DateTime.new(start_time)
  event.dtend = Icalendar::Values::DateTime.new(end_time)
  event.summary = summary
  event.description = description
  event
end

# Process each team
def process_teams(service)
  TEAMS.each do |team|
    process_team(service, team)
  end
end

# Process a single team
def process_team(service, team)
  assignment_counts = initialize_assignment_counts(team)
  calendar = fetch_ical_data(team[:ical_feed_url])
  lrm_calendar = Icalendar::Calendar.new
  csv_data = process_events(calendar, team, assignment_counts, lrm_calendar)
  update_google_sheets(service, team, csv_data)
  save_ical_feed(lrm_calendar, team[:name])
  save_assignment_counts
  save_assigned_events
end

# Initialize assignment counts for a team
def initialize_assignment_counts(team)
  assignment_counts = Hash.new(0)
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }
  assignment_counts
end

# Fetch iCal data
def fetch_ical_data(ical_feed_url)
  uri = URI(ical_feed_url)
  response = Net::HTTP.get(uri)
  Icalendar::Calendar.parse(response).first
end

# Process events for a team
def process_events(calendar, team, assignment_counts, lrm_calendar)
  calendar.events.each_with_index.map do |event, _index|
    next if exclude_event?(event)

    start_time, end_time, duration_in_minutes = calculate_event_times(event)
    next if start_time < Time.now.in_time_zone('Central Time (US & Canada)')

    locker_room_monitor = assign_locker_room_monitor(team, event, assignment_counts)
    create_locker_room_monitor_event(lrm_calendar, start_time, locker_room_monitor, event) if locker_room_monitor
    create_home_game_roles(lrm_calendar, start_time, end_time, event) if home_game?(event)
    prepare_event_data(event, start_time, formatted_date, duration_in_minutes, locker_room_monitor)
  end.compact
end

# Exclude events based on criteria
def exclude_event?(event)
  event.dtstart.nil? || event.dtend.nil? || event.location.nil? || event.location.strip.empty? ||
    EXCLUSION_LIST.any? do |term|
      event.summary&.downcase&.include?(term.downcase) || event.description&.downcase&.include?(term.downcase) || event.location&.downcase&.include?(term.downcase)
    end
end

# Calculate event times and duration
def calculate_event_times(event)
  start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
  end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
  duration_in_minutes = ((end_time - start_time) / 60).to_i
  [start_time, end_time, duration_in_minutes]
end

# Assign locker room monitor
def assign_locker_room_monitor(team, event, assignment_counts)
  @assigned_events[team[:name]][event.uid] || begin
    family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
    assignment_counts[family_with_fewest_assignments] += 1
    @assigned_events[team[:name]][event.uid] = family_with_fewest_assignments
    family_with_fewest_assignments
  end
end

# Create locker room monitor event
def create_locker_room_monitor_event(lrm_calendar, start_time, locker_room_monitor, event)
  lrm_event = Icalendar::Event.new
  lrm_event.dtstart = Icalendar::Values::Date.new(start_time.to_date)
  lrm_event.dtend = Icalendar::Values::Date.new((start_time.to_date + 1))
  lrm_event.summary = locker_room_monitor.force_encoding('UTF-8')
  lrm_event.description = <<-DESC.force_encoding('UTF-8')
    Locker Room Monitor: #{locker_room_monitor}

    Instructions:
    - Locker rooms should be monitored 30 minutes before and closed 15 minutes after the scheduled practice/game.

    Event: #{event.summary.force_encoding('UTF-8')}
    Location: #{event.location.force_encoding('UTF-8')}
    Scheduled Event Time: #{start_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')} to #{end_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')}
  DESC
  lrm_calendar.add_event(lrm_event)
end

# Check if event is a home game
def home_game?(event)
  event.summary.downcase.include?('game') && event.location.downcase.include?('home')
end

# Create additional roles for home games
def create_home_game_roles(lrm_calendar, start_time, end_time, event)
  roles = ['Penalty Box', 'Scorekeeper', 'Timekeeper']
  roles.each do |role|
    role_event = create_ical_event(start_time, end_time, "#{role}: #{event.summary.force_encoding('UTF-8')}",
                                   <<-DESC.force_encoding('UTF-8'))
      #{role} Instructions:
      - Ensure you are ready 30 minutes before the game.
      - Check equipment and uniforms.
      - Coordinate with the coach for any last-minute changes.

      Event: #{event.summary.force_encoding('UTF-8')}
      Location: #{event.location.force_encoding('UTF-8')}
      Scheduled Game Time: #{start_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')} to #{end_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')}
    DESC
    lrm_calendar.add_event(role_event)
  end
end

# Prepare event data for Google Sheets
def prepare_event_data(event, start_time, formatted_date, duration_in_minutes, locker_room_monitor)
  raw_date = start_time.strftime('%Y-%m-%d')
  formatted_date = start_time.strftime('%a %I:%M %p').capitalize
  [event.summary.force_encoding('UTF-8'), event.location.force_encoding('UTF-8'), raw_date, formatted_date,
   duration_in_minutes, locker_room_monitor.force_encoding('UTF-8')]
end

# Update Google Sheets with event data
def update_google_sheets(service, team, csv_data)
  existing_data = fetch_existing_data(service, team[:spreadsheet_id], 'Sheet1!A2:F')
  merged_data = merge_data(existing_data, csv_data)
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, merged_data)
  sheet_id = get_sheet_id(service, team[:spreadsheet_id])
  sort_google_sheet_by_date(service, team[:spreadsheet_id], sheet_id)
end

# Merge existing data with new data, giving priority to Google Sheets data
def merge_data(existing_data, csv_data)
  existing_data_hash = existing_data.to_h { |row| [[row[0], row[1]], row] }
  csv_data.each do |row|
    key = [row[0], row[1]]
    existing_data_hash[key] = row
  end
  existing_data_hash.values
end

# Save iCal feed to a file
def save_ical_feed(lrm_calendar, team_name)
  ics_filename = "locker_room_monitor_#{team_name.downcase.gsub(' ', '_')}.ics"
  lrm_calendar.publish
  File.open(ics_filename, 'w') { |file| file.write(lrm_calendar.to_ical) }
end

# Save assignment counts to file
def save_assignment_counts
  CSV.open(@assignment_counts_file, 'w') do |csv|
    csv << %w[Team Family Count]
    @assignment_counts.sort.each do |team_name, counts|
      counts.each do |family, count|
        csv << [team_name, family, count]
      end
    end
  end
end

# Save assigned events to file
def save_assigned_events
  CSV.open(@assigned_events_file, 'w') do |csv|
    csv << %w[Team EventID Locker_Room_Monitor]
    @assigned_events.each do |team_name, events|
      events.each do |event_id, monitor|
        csv << [team_name, event_id, monitor]
      end
    end
  end
end

# Main execution
service = setup_google_sheets
process_teams(service)

# Display the family counts by team
puts "\nLocker Room Monitor Assignment Counts by Team:"
@assignment_counts.each do |team_name, counts|
  puts "\nTeam #{team_name}:"
  counts.each do |family, count|
    puts "#{family}: #{count}"
  end
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
