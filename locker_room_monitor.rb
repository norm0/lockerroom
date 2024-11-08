require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
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

# Define the authorize method to handle Google OAuth2
def authorize
  client_id = Google::Auth::ClientId.from_file('credentials.json')
  token_store = Google::Auth::Stores::FileTokenStore.new(file: 'token.yaml')
  authorizer = Google::Auth::UserAuthorizer.new(client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(base_url: OOB_URI)
    puts "Open the following URL in your browser and authorize the application:\n#{url}"
    print 'Enter the authorization code: '
    code = gets.chomp
    credentials = authorizer.get_and_store_credentials_from_code(user_id:, code:, base_url: OOB_URI)
  end
  credentials
end

# Initialize the Google Sheets API service
def initialize_service
  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = APPLICATION_NAME
  service.authorization = authorize
  service
end

# Method to clear team data in Google Sheets before updating
def clear_google_sheet_data(service, spreadsheet_id, range)
  clear_request = Google::Apis::SheetsV4::ClearValuesRequest.new
  service.clear_values(spreadsheet_id, range, clear_request)
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
  values = [headers] + # Replace nils with empty strings
           data.map do |row|
                                    row.map do |cell|
                                      cell.nil? ? '' : cell.to_s
                                    end
           end
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
  'Meeting'
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
service = initialize_service

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

  csv_data = calendar.events.each_with_index.map do |event, _index|
    # Skip events if summary or description is nil, empty, or matches exclusion criteria
    next if event.summary.nil? || event.summary.strip.empty? || event.summary.include?('LRM')
    next if event.description.nil? || event.description.strip.empty? || event.description.include?('LRM')
    next if exclusion_list.any? { |term| event.summary.include?(term) || event.description.include?(term) }
    next if event.dtstart.nil? || event.dtend.nil?

    # Process the event if it’s not excluded
    event_id = event.uid
    start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
    end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
    raw_date = start_time.strftime('%Y-%m-%d')
    formatted_date = start_time.strftime('%a %I:%M %p').downcase
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Balanced random assignment of locker room monitor per team
    locker_room_monitor = @assigned_events[event_id] || begin
      family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
      assignment_counts[family_with_fewest_assignments] += 1
      @assigned_events[event_id] = family_with_fewest_assignments
      family_with_fewest_assignments
    end

    # Prepare data for Google Sheets
    [event.summary, event.location, raw_date, formatted_date, duration_in_minutes, locker_room_monitor]
  end.compact

  # Clear existing data and write new data to Google Sheets
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, csv_data)

  # Save assignment counts to ensure persistence
  CSV.open("#{@assignment_counts_file}_#{team[:name]}.csv", 'w') do |csv|
    csv << %w[Family Count]
    assignment_counts.each { |family, count| csv << [family, count] }
  end

  # Save assigned events to ensure persistence
  CSV.open("#{@assigned_events_file}_#{team[:name]}.csv", 'w') do |csv|
    csv << %w[EventID Locker_Room_Monitor]
    assigned_events.each do |event_id, monitor|
      csv << [event_id, monitor]
    end
  end
end

# Display the family counts by team
puts "\nLocker Room Monitor Assignment Counts by Team:"
teams.each do |team|
  puts "\nTeam #{team[:name]}:"
  team[:family_names].each do |family|
    puts "#{family}: #{assignment_counts[family]}"
  end
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
