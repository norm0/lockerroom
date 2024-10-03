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

# Initialize a new iCal feed for locker room monitor events
lrm_calendar = Icalendar::Calendar.new

# Assign a locker room monitor for each event, cycling through families
csv_data = filtered_events.each_with_index.map do |event, index|
  # Cycle through the family_names array
  locker_room_monitor = family_names[index % family_names.size]

  # Create an all-day event with instructions for the locker room monitor
  lrm_event = Icalendar::Event.new
  lrm_event.dtstart = Icalendar::Values::Date.new(event.dtstart.to_date) # All-day event starts on the event date
  lrm_event.dtend = Icalendar::Values::Date.new((event.dtstart + 1.day).to_date) # End date is the next day (to mark all-day event)
  lrm_event.summary = "LRM #{locker_room_monitor}"
  lrm_event.description = <<-DESC
    Locker Room Monitor: #{locker_room_monitor}

    Instructions:
    - Locker rooms should be opened 30 minutes before the scheduled practice/game.
    - Locker rooms should be monitored and closed 15 minutes after the scheduled practice/game.

    Event: #{event.summary}
    Location: #{event.location}
    Scheduled Event Time: #{event.dtstart.strftime('%a, %b %-d, %Y at %-I:%M %p')} to #{event.dtend.strftime('%a, %b %-d, %Y at %-I:%M %p')}
  DESC

  # Add the event to the iCal feed
  lrm_calendar.add_event(lrm_event)

  # Return data for CSV export
  {
    'Event' => event.summary,
    'Location' => event.location,
    'Start Date' => event.dtstart.strftime('%a, %b %-d, %Y'),
    'End Date' => event.dtend.strftime('%a, %b %-d, %Y'),
    'Locker Room Monitor' => locker_room_monitor
  }
end

# Save CSV file
CSV.open('locker_room_monitors.csv', 'w') do |csv|
  csv << ['Event', 'Location', 'Start Date', 'End Date', 'Locker Room Monitor']
  csv_data.each do |row|
    csv << [row['Event'], row['Location'], row['Start Date'], row['End Date'], row['Locker Room Monitor']]
  end
end

# Finalize the iCal feed and save it
lrm_calendar.publish
File.open('locker_room_monitor.ics', 'w') { |file| file.write(lrm_calendar.to_ical) }

puts "Events with locker room monitors have been saved to 'locker_room_monitors.csv' and the iCal feed has been saved to 'locker_room_monitor.ics'."
