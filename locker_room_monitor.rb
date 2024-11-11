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

@assignment_counts_file = 'assignment_counts.csv'
@assigned_events_file = 'assigned_events.csv'

@assignment_counts = Hash.new { |hash, key| hash[key] = Hash.new(0) }
@assigned_events = Hash.new { |hash, key| hash[key] = {} }

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

load_assignment_counts
load_assigned_events

def setup_google_sheets
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize
  service
end

def authorize
  credentials = JSON.parse(ENV['GOOGLE_SHEETS_CREDENTIALS'])
  Google::Auth::ServiceAccountCredentials.make_creds(
    json_key_io: StringIO.new(credentials.to_json),
    scope: SCOPE
  )
end

def clear_google_sheet_data(service, spreadsheet_id, range)
  clear_request = Google::Apis::SheetsV4::ClearValuesRequest.new
  service.clear_values(spreadsheet_id, range, clear_request)
end

def fetch_existing_data(service, spreadsheet_id, range)
  response = service.get_spreadsheet_values(spreadsheet_id, range)
  response.values || []
end

def sort_google_sheet_by_date(service, spreadsheet_id, sheet_id)
  sort_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
    requests: [
      {
        sort_range: {
          range: {
            sheet_id:,
            start_row_index: 1,
            start_column_index: 0,
            end_column_index: 9 # Adjusted for new columns
          },
          sort_specs: [
            {
              dimension_index: 2,
              sort_order: 'ASCENDING'
            }
          ]
        }
      }
    ]
  )
  service.batch_update_spreadsheet(spreadsheet_id, sort_request)
end

def write_team_data_to_individual_sheets(service, team, data)
  headers = ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor', 'Penalty Box',
             'Scorekeeper', 'Timekeeper']
  values = [headers] + data
  range = 'Sheet1!A1:I' # Updated range to accommodate new columns
  value_range = Google::Apis::SheetsV4::ValueRange.new(values:)
  service.update_spreadsheet_value(team[:spreadsheet_id], range, value_range, value_input_option: 'RAW')
end

EXCLUSION_LIST = [
  'Skills Off Ice', 'Dryland', 'Goalie Training', 'Off Ice', 'Conditioning', 'Meeting', 'Goalie', 'LRM', 'tournament', 'Tournament', 'Pictures'
]

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

def create_ical_event(start_time, end_time, summary, description)
  event = Icalendar::Event.new
  event.dtstart = Icalendar::Values::DateTime.new(start_time)
  event.dtend = Icalendar::Values::DateTime.new(end_time)
  event.summary = summary
  event.description = description
  event
end

def process_teams(service)
  TEAMS.each do |team|
    process_team(service, team)
  end
end

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

def initialize_assignment_counts(team)
  assignment_counts = Hash.new(0)
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }
  assignment_counts
end

def fetch_ical_data(ical_feed_url)
  uri = URI(ical_feed_url)
  response = Net::HTTP.get(uri)
  Icalendar::Calendar.parse(response).first
end

def process_events(calendar, team, assignment_counts, lrm_calendar)
  calendar.events.each_with_index.map do |event, _index|
    next if exclude_event?(event)

    start_time, end_time, duration_in_minutes = calculate_event_times(event)
    next if start_time < Time.now.in_time_zone('Central Time (US & Canada)')

    locker_room_monitor = assign_locker_room_monitor(team, event, assignment_counts)
    create_locker_room_monitor_event(lrm_calendar, start_time, locker_room_monitor, event) if locker_room_monitor

    # Assign additional roles if it's a home game
    penalty_box, scorekeeper, timekeeper = assign_additional_roles(event, team, assignment_counts)

    prepare_event_data(event, start_time, duration_in_minutes, locker_room_monitor, penalty_box, scorekeeper,
                       timekeeper)
  end.compact
end

def exclude_event?(event)
  event.dtstart.nil? || event.dtend.nil? || event.location.nil? || event.location.strip.empty? ||
    EXCLUSION_LIST.any? do |term|
      event.summary&.downcase&.include?(term.downcase) || event.description&.downcase&.include?(term.downcase) || event.location&.downcase&.include?(term.downcase)
    end
end

def calculate_event_times(event)
  start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
  end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
  duration_in_minutes = ((end_time - start_time) / 60).to_i
  [start_time, end_time, duration_in_minutes]
end

def assign_locker_room_monitor(team, event, assignment_counts)
  @assigned_events[team[:name]][event.uid] || begin
    family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
    assignment_counts[family_with_fewest_assignments] += 1
    @assigned_events[team[:name]][event.uid] = family_with_fewest_assignments
    family_with_fewest_assignments
  end
end

# Assign additional roles for home games
def assign_additional_roles(event, team, assignment_counts)
  if event.summary.downcase.include?('game') && event.location.downcase.include?('home')
    %w[Penalty Box Scorekeeper Timekeeper].map do
      family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
      assignment_counts[family_with_fewest_assignments] += 1
      family_with_fewest_assignments
    end
  else
    [nil, nil, nil]
  end
end

def prepare_event_data(event, start_time, duration_in_minutes, locker_room_monitor, penalty_box, scorekeeper,
                       timekeeper)
  raw_date = start_time.strftime('%Y-%m-%d')
  formatted_date = start_time.strftime('%a %I:%M %p').capitalize
  [event.summary.force_encoding('UTF-8'), event.location.force_encoding('UTF-8'), raw_date, formatted_date,
   duration_in_minutes, locker_room_monitor.force_encoding('UTF-8'), penalty_box, scorekeeper, timekeeper]
end

def update_google_sheets(service, team, csv_data)
  existing_data = fetch_existing_data(service, team[:spreadsheet_id], 'Sheet1!A2:I')
  merged_data = merge_data(existing_data, csv_data)
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:I')
  write_team_data_to_individual_sheets(service, team, merged_data)
  sheet_id = get_sheet_id(service, team[:spreadsheet_id])
  sort_google_sheet_by_date(service, team[:spreadsheet_id], sheet_id)
end

def merge_data(existing_data, csv_data)
  existing_data_hash = existing_data.to_h { |row| [[row[0], row[1]], row] }
  csv_data.each do |row|
    key = [row[0], row[1]]
    existing_data_hash[key] = row
  end
  existing_data_hash.values
end

def save_ical_feed(lrm_calendar, team_name)
  ics_filename = "locker_room_monitor_#{team_name.downcase.gsub(' ', '_')}.ics"
  lrm_calendar.publish
  File.open(ics_filename, 'w') { |file| file.write(lrm_calendar.to_ical) }
end

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

service = setup_google_sheets
process_teams(service)

puts "\nLocker Room Monitor Assignment Counts by Team:"
@assignment_counts.each do |team_name, counts|
  puts "\nTeam #{team_name}:"
  counts.each do |family, count|
    puts "#{family}: #{count}"
  end
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
