require 'net/http'
require 'icalendar'
require 'uri'
require 'csv'

# File to store assignment counts and assigned events
assignment_counts_file = 'assignment_counts.csv'
assigned_events_file = 'assigned_events.csv'

# Read existing assignment counts from CSV file
assignment_counts = Hash.new(0)
if File.exist?(assignment_counts_file)
  CSV.foreach(assignment_counts_file, headers: true) do |row|
    assignment_counts[row['Family']] = row['Count'].to_i
  end
end

# Read existing assigned events from CSV file
assigned_events = {}
if File.exist?(assigned_events_file)
  CSV.foreach(assigned_events_file, headers: true) do |row|
    assigned_events[row['EventID']] = row['Locker Room Monitor']
  end
end

# Team configurations for 12A, 12B1, 10B1, and 10B2
teams = [
  {
    name: '12A',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603019',
    family_names: %w[Becker Hastings Opel Gorgos Larsen Anderson Campos Powell Tousignant Marshall Johnson Wulff Orstad
                     Mulcahey]
  },
  {
    name: '12B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603021',
    family_names: %w[Baer Bimberg Chanthavongsa Hammerstrom Kremer Lane Oas Perpich Ray Reinke Silva-Hammer]
  },
  {
    name: '10B1',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603022',
    family_names: %w[Baer Bowman Hopper Houghtaling Johnson Larsen Markfort Marshall Nanninga Orstad Willey Williamson]
  },
  {
    name: '10B2',
    ical_feed_url: 'https://www.armstrongcooperhockey.org/ical_feed?tags=8603023',
    family_names: %w[Curry Engholm Froberg Harpel Johnson Mckinnon Oprenchak M-Reberg B-Reberg Sauer Smith Woods]
  }
]

# Locations that require locker room monitors
locations_with_monitors = ['New Hope North', 'New Hope South', 'Breck', 'Orono Ice Arena (ag)', 'Northeast (ag)',
                           'SLP East (ag)', 'MG West (ag)', 'PIC A (ag)', 'PIC C (ag)', 'Hopkins Pavilion (ag)', 'Thaler (ag)', 'SLP West (ag)', 'Delano Arena']

# Track the last assigned family for each team to avoid back-to-back assignments
last_assigned_family = {}

teams.each do |team|
  # Initialize the assignment count for each family in the team if not already present
  team[:family_names].each { |family| assignment_counts[family] ||= 0 }

  # Fetch the iCal feed using Net::HTTP for the current team
  uri = URI(team[:ical_feed_url])
  response = Net::HTTP.get(uri)

  # Parse the iCal feed
  calendar = Icalendar::Calendar.parse(response).first

  # Initialize a new iCal feed for locker room monitor events
  lrm_calendar = Icalendar::Calendar.new

  # Assign a locker room monitor for each event, cycling through family names
  csv_data = calendar.events.each_with_index.map do |event, _index|
    # Generate a unique event ID (e.g., using the event UID)
    event_id = event.uid

    # Skip events without a start or end time
    next if event.dtstart.nil? || event.dtend.nil?

    # Skip all-day events with "LRM" in the summary or description
    if event.dtstart.is_a?(Icalendar::Values::Date) && (event.summary&.include?('LRM') || event.description&.include?('LRM'))
      next
    end

    # Convert Icalendar::Values::DateTime to Ruby Time object
    start_time = event.dtstart.to_time
    end_time = event.dtend.to_time

    # Format the date and time separately (human-readable)
    date_formatted = start_time.strftime('%Y-%m-%d')
    time_formatted = start_time.strftime('%I:%M %p') # 12-hour format with AM/PM

    # Calculate duration in minutes for non-all-day events
    duration_in_minutes = ((end_time - start_time) / 60).to_i

    # Determine if this event location requires a locker room monitor
    locker_room_monitor = if locations_with_monitors.include?(event.location)
                            # Check if the event has already been assigned a monitor
                            assigned_events[event_id] || begin
                              # Shuffle the family names to randomize the order
                              shuffled_families = team[:family_names].shuffle
                              # Select the family with the fewest assignments, avoiding back-to-back assignments
                              eligible_families = shuffled_families.reject do |family|
                                family == last_assigned_family[team[:name]]
                              end
                              family_with_fewest_assignments = eligible_families.min_by do |family|
                                assignment_counts[family]
                              end
                              # Update the count for the selected family
                              assignment_counts[family_with_fewest_assignments] += 1
                              # Track the assignment
                              assigned_events[event_id] = family_with_fewest_assignments
                              # Update the last assigned family for the team
                              last_assigned_family[team[:name]] = family_with_fewest_assignments
                              family_with_fewest_assignments
                            end
                          else
                            nil # No locker room monitor for other locations
                          end

    # Create an all-day event with instructions for the locker room monitor (only if required)
    if locker_room_monitor
      lrm_event = Icalendar::Event.new
      lrm_event.dtstart = Icalendar::Values::Date.new(start_time.to_date) # All-day event starts on the event date
      lrm_event.dtend = Icalendar::Values::Date.new((start_time.to_date + 1)) # End date is the next day (to mark all-day event)
      lrm_event.summary = "#{locker_room_monitor}" # Only the monitor's name
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

      # Set duration to nil for all-day events
      duration_in_minutes = nil
    end

    # Return data for CSV export
    {
      'Event' => event.summary,
      'Location' => event.location,
      'Date' => date_formatted,
      'Time' => time_formatted,
      'Duration (minutes)' => duration_in_minutes,
      'Locker Room Monitor' => locker_room_monitor || ''
    }
  end.compact # Remove nil values from the array

  # Sort the CSV data by 'Start Time' before writing to the CSV file
  csv_data.sort_by! { |row| row['Date'] + row['Time'] }

  # Save CSV file with a team-specific filename
  csv_filename = "locker_room_monitors_#{team[:name].downcase.gsub(' ', '_')}.csv"
  CSV.open(csv_filename, 'w') do |csv|
    csv << ['Event', 'Location', 'Date', 'Time', 'Duration (minutes)', 'Locker Room Monitor']
    csv_data.each do |row|
      csv << [row['Event'], row['Location'], row['Date'], row['Time'], row['Duration (minutes)'],
              row['Locker Room Monitor']]
    end
  end

  # Finalize the iCal feed and save it with a team-specific filename
  ics_filename = "locker_room_monitor_#{team[:name].downcase.gsub(' ', '_')}.ics"
  lrm_calendar.publish
  File.open(ics_filename, 'w') { |file| file.write(lrm_calendar.to_ical) }

  puts "Events with locker room monitors for #{team[:name]} have been saved to '#{csv_filename}' and the iCal feed has been saved to '#{ics_filename}'."
end

# Write updated assignment counts back to the CSV file
CSV.open(assignment_counts_file, 'w') do |csv|
  csv << %w[Family Count]
  assignment_counts.each do |family, count|
    csv << [family, count]
  end
end

# Write updated assigned events back to the CSV file
CSV.open(assigned_events_file, 'w') do |csv|
  csv << ['EventID', 'Locker Room Monitor']
  assigned_events.each do |event_id, monitor|
    csv << [event_id, monitor]
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
