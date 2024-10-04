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

  # Convert Icalendar::Values::DateTime to Ruby Time object
  start_time = event.dtstart.to_time
  end_time = event.dtend.to_time

  # Format the date and time separately
  date_formatted = start_time.strftime('%Y-%m-%d')
  time_formatted = start_time.strftime('%H:%M:%S')

  # Calculate duration in minutes
  duration_in_minutes = ((end_time - start_time) / 60).to_i

  # Create an all-day event with instructions for the locker room monitor
  lrm_event = Icalendar::Event.new
  lrm_event.dtstart = Icalendar::Values::Date.new(start_time.to_date) # All-day event starts on the event date
  lrm_event.dtend = Icalendar::Values::Date.new((start_time.to_date + 1)) # End date is the next day (to mark all-day event)
  lrm_event.summary = "LRM #{locker_room_monitor}"
  lrm_event.description = <<-DESC
    Locker Room Monitor: #{locker_room_monitor}

    Instructions:
    - Locker rooms should be monitored 30 minutes before and closed 15 minutes after the scheduled practice/game.

    Event: #{event.summary}
    Location: #{event.location}
    Scheduled Event Time: #{start_time.strftime('%a, %b %-d, %Y at %-I:%M %p')} to #{end_time.strftime('%a, %b %-d, %Y at %-I:%M %p')}
  DESC

  # Add the event to the iCal feed
  lrm_calendar.add_event(lrm_event)

  # Return data for CSV export
  {
    'Event' => event.summary,
    'Location' => event.location,
    'Date' => date_formatted,
    'Time' => time_formatted,
    'Duration (minutes)' => duration_in_minutes,
    'Locker Room Monitor' => locker_room_monitor
  }
end

# Save CSV file with lowercase filename
CSV.open('locker_room_monitors.csv', 'w') do |csv|
  csv << ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
  csv_data.each do |row|
    csv << [row['Event'], row['Location'], row['Date'], row['Time'], row['Duration (minutes)'],
            row['Locker Room Monitor']]
  end
end

# Finalize the iCal feed and save it with a lowercase filename
lrm_calendar.publish
File.open('locker_room_monitor.ics', 'w') { |file| file.write(lrm_calendar.to_ical) }

puts "Events with locker room monitors have been saved to 'locker_room_monitors.csv' and the iCal feed has been saved to 'locker_room_monitor.ics'."
