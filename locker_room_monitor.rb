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
SCOPE = [Google::Apis::SheetsV4::AUTH_SPREADSHEETS]

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

# Fetch data from Google Sheets and merge with local CSVs
def fetch_and_merge_google_sheet_data(service, team)
  sheet_data = service.get_spreadsheet_values(team[:spreadsheet_id], 'Sheet1!A2:F').values

  # Update assigned events and assignment counts from Google Sheets
  sheet_data.each do |row|
    event_id = row[0]
    monitor = row[5] # Assuming event_id and monitor columns in the Sheet

    # If monitor data is missing, skip
    next unless monitor

    # Check if we need to update or add this event in assigned_events
    if @assigned_events[event_id] != monitor
      @assigned_events[event_id] = monitor
      @assignment_counts[monitor] += 1 unless monitor.empty?
    end
  end

  # Save merged data to CSV files
  save_assignment_counts
  save_assigned_events
end

# Save assignment counts to CSV
def save_assignment_counts
  CSV.open(@assignment_counts_file, 'w') do |csv|
    csv << %w[Family Count]
    @assignment_counts.each { |family, count| csv << [family, count] }
  end
end

# Save assigned events to CSV
def save_assigned_events
  CSV.open(@assigned_events_file, 'w') do |csv|
    csv << ['EventID', 'Locker Room Monitor']
    @assigned_events.each { |event_id, monitor| csv << [event_id, monitor] }
  end
end

# Method to write team data to Google Sheets
def write_team_data_to_individual_sheets(service, team, data)
  headers = ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
  values = [headers] + data
  range = 'Sheet1!A1:F'
  value_range = Google::Apis::SheetsV4::ValueRange.new(values:)
  service.update_spreadsheet_value(team[:spreadsheet_id], range, value_range, value_input_option: 'RAW')
end

# Team configurations for each team
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

# Locations that require locker room monitors
locations_with_monitors = ['New Hope North', 'New Hope South', 'Breck', 'Orono Ice Arena (ag)', 'Northeast (ag)',
                           'SLP East (ag)', 'MG West (ag)', 'PIC A (ag)', 'PIC C (ag)', 'Hopkins Pavilion (ag)', 'Thaler (ag)', 'SLP West (ag)', 'Delano Arena']

# Locations that do not require locker room monitors
excluded_locations = [
  'New Hope North - Skills Off Ice',
  'New Hope Ice Arena, Louisiana Avenue North, New Hope, MN, USA',
  nil, '' # Empty locations
]

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

def get_sheet_id(service, spreadsheet_id)
  spreadsheet = service.get_spreadsheet(spreadsheet_id)
  spreadsheet.sheets.first.properties.sheet_id # Assumes only one sheet
end

# Fetch, merge, and update data for each team
service = setup_google_sheets
teams.each do |team|
  fetch_and_merge_google_sheet_data(service, team)

  # Initialize iCal calendar for each team
  lrm_calendar = Icalendar::Calendar.new

  # Fetch iCal data, process events, and update Google Sheets
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)
  calendar = Icalendar::Calendar.parse(response).first

  csv_data = calendar.events.each_with_index.map do |event, index|
    # Skip if the event contains 'LRM' in the summary or description
    next if event.summary&.include?('LRM') || event.description&.include?('LRM')

    # Skip if the location is in the excluded locations
    location = event.location
    next if excluded_locations.include?(location)

    # Skip if event has no start or end time
    next if event.dtstart.nil? || event.dtend.nil?

    # Process event times
    event_id = event.uid
    start_time = event.dtstart.to_time.in_time_zone('Central Time (US & Canada)')
    end_time = event.dtend.to_time.in_time_zone('Central Time (US & Canada)')
    date_formatted = start_time.strftime('%A %m/%d')
    time_formatted = start_time.strftime('%I:%M %p %Z')
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Determine if this event location requires a locker room monitor
    locker_room_monitor = if locations_with_monitors.include?(location)
                            @assigned_events[event_id] || team[:family_names][index % team[:family_names].size]
                          end

    # Prepare data for Google Sheets only if it passed all exclusion checks
    [event.summary, location, date_formatted, time_formatted, duration_in_minutes, locker_room_monitor]
  end.compact

  # Clear existing data and write new data to Google Sheets
  clear_google_sheet_data(service, team[:spreadsheet_id], 'Sheet1!A1:F')
  write_team_data_to_individual_sheets(service, team, csv_data)

  # Sort the sheet by date
  sheet_id = get_sheet_id(service, team[:spreadsheet_id]) # Get the sheet ID dynamically
  sort_google_sheet_by_date(service, team[:spreadsheet_id], sheet_id)

  # Write iCal file for each team
  ics_filename = "locker_room_monitor_#{team[:name].downcase.gsub(' ', '_')}.ics"
  File.open(ics_filename, 'w') { |file| file.write(lrm_calendar.to_ical) }
end

puts 'Data fetched, merged, and updated successfully, including .ics files.'
