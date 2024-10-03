require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'

# iCal feed URL (using https)
ical_feed_url = 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019'

# Family names provided
family_names = %w[
  Becker Hastings Opel Gorgos Larsen
  Anderson Orstad Campos Powell Tousignant Marshall Johnson Wulff
]

# Fetch the iCal feed using Net::HTTP
uri = URI(ical_feed_url)
response = Net::HTTP.get(uri)

# Parse the iCal feed
calendar = Icalendar::Calendar.parse(response).first

# Filter for events that have a start and end time and are at specific locations
filtered_events = calendar.events.select do |event|
  event.dtstart && event.dtend &&
    ['New Hope North', 'New Hope South', 'Breck'].include?(event.location)
end

# Helper method to format dates with abbreviated day and month names
def format_friendly_date(datetime)
  datetime.strftime('%a, %b %-d, %Y at %-I:%M %p')
end

# Initialize a new iCal feed for locker room monitor events
lrm_calendar = Icalendar::Calendar.new

# Assign a locker room monitor for each event, cycling through families
csv_data = filtered_events.each_with_index.map do |event, index|
  # Cycle through the family_names array
  locker_room_monitor = family_names[index % family_names.size]

  # Create two new LRM events: one 30 minutes before and one 15 minutes after the event
  lrm_before = event.dtstart - (30 * 60) # 30 minutes before
  lrm_after = event.dtend + (15 * 60)    # 15 minutes after

  # Add LRM events to the iCal feed
  lrm_event_before = Icalendar::Event.new
  lrm_event_before.dtstart = lrm_before
  lrm_event_before.dtend = lrm_before + (15 * 60) # Duration of 15 minutes
  lrm_event_before.summary = "LRM #{locker_room_monitor}"
  lrm_calendar.add_event(lrm_event_before)

  lrm_event_after = Icalendar::Event.new
  lrm_event_after.dtstart = lrm_after
  lrm_event_after.dtend = lrm_after + (15 * 60) # Duration of 15 minutes
  lrm_event_after.summary = "LRM #{locker_room_monitor}"
  lrm_calendar.add_event(lrm_event_after)

  # Return data for CSV export
  {
    'Event' => event.summary,
    'Location' => event.location,
    'Start Date' => format_friendly_date(event.dtstart),
    'End Date' => format_friendly_date(event.dtend),
    'Locker Room Monitor' => locker_room_monitor
  }
end

# Print the events with the assigned locker room monitors and friendly dates
csv_data.each do |row|
  puts "Event: #{row['Event']}"
  puts "Location: #{row['Location']}"
  puts "Start Date: #{row['Start Date']}"
  puts "End Date: #{row['End Date']}"
  puts "Locker Room Monitor: #{row['Locker Room Monitor']}"
  puts '-' * 40
end

# Save CSV file
CSV.open('Locker_Room_Monitors.csv', 'w') do |csv|
  csv << ['Event', 'Location', 'Start Date', 'End Date', 'Locker Room Monitor']
  csv_data.each do |row|
    csv << [row['Event'], row['Location'], row['Start Date'], row['End Date'], row['Locker Room Monitor']]
  end
end

# Finalize the iCal feed and save it
lrm_calendar.publish
File.open('Locker_Room_Monitor.ics', 'w') { |file| file.write(lrm_calendar.to_ical) }

puts "Events with locker room monitors have been saved to 'Locker_Room_Monitors.csv' and the iCal feed has been saved to 'Locker_Room_Monitor.ics'."
