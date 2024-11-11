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
if File.exist?(@assignment_counts_file)
  CSV.foreach(@assignment_counts_file, headers: true) do |row|
    team_name = row['Team']
    family = row['Family']
    count = row['Count'].to_i
    @assignment_counts[team_name][family] = count
  end
end

if File.exist?(@assigned_events_file)
  CSV.foreach(@assigned_events_file, headers: true) do |row|
    team_name = row['Team']
    event_id = row['EventID']
    monitor = row['Locker Room Monitor']
    @assigned_events[team_name][event_id] = monitor
  end
end

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

  # Set up ServiceAccountCredentials using the JSON key
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
exclusion_list = [
  'Skills Off Ice', # Example keywords or patterns
  'Dryland',
  'Goalie Training',
  'Off Ice',
  'Conditioning',
  'Meeting',
  'Goalie',
  'LRM',
  'tournament',
  'Tournament',
  'picture'
]

# Define teams and configurations
teams = [
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

# Fetch, merge, and update data for each team
service = setup_google_sheets

# Method to retrieve the sheet ID for sorting
def get_sheet_id(service, spreadsheet_id)
  spreadsheet = service.get_spreadsheet(spreadsheet_id)
  spreadsheet.sheets.first.properties.sheet_id # Assumes only one sheet
end

teams.each do |team|
  # Initialize a separate assignment count for each team
  assignment_counts = Hash.new(0)
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }

  # Fetch iCal data, process events, and update Google Sheets
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)
  calendar = Icalendar::Calendar.parse(response).first

  # Initialize a new iCal feed for locker room monitor events
  lrm_calendar = Icalendar::Calendar.new

  csv_data = calendar.events.each_with_index.map do |event, _index|
    # Skip events if summary, description, or location matches exclusion criteria
    if exclusion_list.any? do |term|
         event.summary&.downcase&.include?(term.downcase) || event.description&.downcase&.include?(term.downcase) || event.location&.downcase&.include?(term.downcase)
       end
      puts "Excluding event: #{event.summary} at #{event.location}"
      next
    end
    next if event.dtstart.nil? || event.dtend.nil?
    next if event.location.nil? || event.location.strip.empty?

    # Filter events to include only those starting from today onwards
    start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
    next if start_time < Time.now.in_time_zone('Central Time (US & Canada)')

    end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
    raw_date = start_time.strftime('%Y-%m-%d')
    formatted_date = start_time.strftime('%a %I:%M %p').capitalize
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Debugging output for duration
    puts "Event: #{event.summary}, Start Time: #{start_time}, End Time: #{end_time}, Duration: #{duration_in_minutes} minutes"

    # Balanced random assignment of locker room monitor per team
    locker_room_monitor = @assigned_events[team[:name]][event.uid] || begin
      family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
      assignment_counts[family_with_fewest_assignments] += 1
      @assigned_events[team[:name]][event.uid] = family_with_fewest_assignments
      family_with_fewest_assignments
    end

    # Create an all-day event with instructions for the locker room monitor (only if required)
    if locker_room_monitor
      lrm_event = Icalendar::Event.new
      lrm_event.dtstart = Icalendar::Values::Date.new(start_time.to_date) # All-day event starts on the event date
      lrm_event.dtend = Icalendar::Values::Date.new((start_time.to_date + 1)) # End date is the next day (to mark all-day event)
      lrm_event.summary = locker_room_monitor.force_encoding('UTF-8') # Only the monitor's name
      lrm_event.description = <<-DESC.force_encoding('UTF-8')
        Locker Room Monitor: #{locker_room_monitor}

        Instructions:
        - Locker rooms should be monitored 30 minutes before and closed 15 minutes after the scheduled practice/game.

        Event: #{event.summary.force_encoding('UTF-8')}
        Location: #{event.location.force_encoding('UTF-8')}
        Scheduled Event Time: #{start_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')} to #{end_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')}
      DESC

      # Add the event to the iCal feed
      lrm_calendar.add_event(lrm_event)

      # Set duration to nil for all-day events
      duration_in_minutes = nil
    end

    # Create additional roles for home games
    if event.summary.downcase.include?('game') && event.location.downcase.include?('home')
      roles = ['Penalty Box', 'Scorekeeper', 'Timekeeper']
      roles.each do |role|
        role_event = Icalendar::Event.new
        role_event.dtstart = Icalendar::Values::DateTime.new(start_time) # Use the original start time
        role_event.dtend = Icalendar::Values::DateTime.new(end_time) # Use the original end time
        role_event.summary = "#{role}: #{event.summary.force_encoding('UTF-8')}"
        role_event.description = <<-DESC.force_encoding('UTF-8')
          #{role} Instructions:
          - Ensure you are ready 30 minutes before the game.
          - Check equipment and uniforms.
          - Coordinate with the coach for any last-minute changes.

          Event: #{event.summary.force_encoding('UTF-8')}
          Location: #{event.location.force_encoding('UTF-8')}
          Scheduled Game Time: #{start_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')} to #{end_time.strftime('%a, %b %-d, %Y at %-I:%M %p').force_encoding('UTF-8')}
        DESC

        # Add the role event to the iCal feed
        lrm_calendar.add_event(role_event)
      end
    end

    # Prepare data for Google Sheets
    [event.summary.force_encoding('UTF-8'), event.location.force_encoding('UTF-8'), raw_date, formatted_date,
     duration_in_minutes, locker_room_monitor.force_encoding('UTF-8')]
  end.compact

  # Fetch existing data from Google Sheets
  existing_data = fetch_existing_data(service, team[:spreadsheet_id], 'Sheet1!A2:F') # Skip header row

  # Merge existing data with new data, giving priority to Google Sheets data
  existing_data_hash = existing_data.to_h { |row| [[row[0], row[1]], row] }
  csv_data.each do |row|
    key = [row[0], row[1]]
    existing_data_hash[key] = row
  end
  merged_data = existing_data_hash.values

  # Clear existing data and write merged data to Google Sheets
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, merged_data)

  # Retrieve the sheet ID and sort by date
  sheet_id = get_sheet_id(service, team[:spreadsheet_id])
  sort_google_sheet_by_date(service, team[:spreadsheet_id], sheet_id)

  # Save assignment counts to ensure persistence
  @assignment_counts[team[:name]] = assignment_counts
  CSV.open(@assignment_counts_file, 'w') do |csv|
    csv << %w[Team Family Count]
    @assignment_counts.sort.each do |team_name, counts|
      counts.each do |family, count|
        csv << [team_name, family, count]
      end
    end
  end

  # Save assigned events to ensure persistence
  CSV.open(@assigned_events_file, 'w') do |csv|
    csv << %w[Team EventID Locker_Room_Monitor]
    @assigned_events.each do |team_name, events|
      events.each do |event_id, monitor|
        csv << [team_name, event_id, monitor]
      end
    end
  end

  # Save the iCal feed to a file
  ics_filename = "locker_room_monitor_#{team[:name].downcase.gsub(' ', '_')}.ics"
  lrm_calendar.publish
  File.open(ics_filename, 'w') { |file| file.write(lrm_calendar.to_ical) }

  puts "Events with locker room monitors for #{team[:name]} have been saved to '#{ics_filename}' and the iCal feed has been saved to '#{ics_filename}'."
end

# Display the family counts by team
puts "\nLocker Room Monitor Assignment Counts by Team:"
@assignment_counts.each do |team_name, counts|
  puts "\nTeam #{team_name}:"
  counts.each do |family, count|
    puts "#{family}: #{count}"
  end
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
