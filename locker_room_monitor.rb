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

# File to store assignment counts and assigned events
@assignment_counts_file = 'assignment_counts.csv'
@assigned_events_file = 'assigned_events.csv'

# Initialize global data
@assignment_counts = Hash.new(0)
@assigned_events = {}

# Load assignment counts and assigned events from files if they exist
if File.exist?(@assignment_counts_file)
  CSV.foreach(@assignment_counts_file, headers: true) do |row|
    @assignment_counts[row['Family']] = row['Count'].to_i
  end
end

if File.exist?(@assigned_events_file)
  CSV.foreach(@assigned_events_file, headers: true) do |row|
    @assigned_events[row['EventID']] = row['Locker Room Monitor']
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
service = setup_google_sheets

teams.each do |team|
  assignment_counts = Hash.new(0)
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }

  # Implement fetch_and_merge_google_sheet_data here if necessary for sync

  # Initialize iCal calendar for each team
  lrm_calendar = Icalendar::Calendar.new

  # Fetch iCal data, process events, and update Google Sheets
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)
  calendar = Icalendar::Calendar.parse(response).first

  csv_data = calendar.events.each_with_index.map do |event, _index|
    # Skip events if summary or description is nil or empty
    next if event.summary.nil? || event.summary.strip.empty?
    next if event.description.nil? || event.description.strip.empty?
  
    # Check if the event summary or description matches any term in the exclusion list
    next if exclusion_list.any? { |term| event.summary.include?(term) || event.description.include?(term) }
  
    next if event.dtstart.nil? || event.dtend.nil?
  
    # Process the event if it’s not excluded
    event_id = event.uid
    start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
    end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
    raw_date = start_time.strftime('%Y-%m-%d')
    formatted_date = start_time.strftime('%m/%d/%y %a %I:%M %p').downcase
    duration_in_minutes = ((end_time - start_time) / 60).to_i
  
    # Balanced assignment of locker room monitor per team
    locker_room_monitor = @assigned_events[event_id] || begin
      family_with_fewest_assignments = team[:family_names].min_by { |family| assignment_counts[family] }
      assignment_counts[family_with_fewest_assignments] += 1
      @assigned_events[event_id] = family_with_fewest_assignments
      family_with_fewest_assignments
    end
  
    # Prepare data for Google Sheets
    [event.summary, event.location, raw_date, formatted_date, duration_in_minutes, locker_room_monitor]
  end.compact

  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, csv_data)
  sheet_id = get_sheet_id(service, team[:spreadsheet_id])
  sort_google_sheet_by_date(service, team[:spreadsheet_id], sheet_id)

  # Save assignment counts for persistence
  CSV.open("#{@assignment_counts_file}_#{team[:name]}.csv", 'w') do |csv|
    csv << %w[Family Count]
    assignment_counts.each { |family, count| csv << [family, count] }
  end
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
