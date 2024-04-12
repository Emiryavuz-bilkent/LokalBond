function addBookingToCalendar() {
    var label = GmailApp.getUserLabelByName("Activity Booking");
    var threads = label.getThreads();
    
    // Duration in hours for each product code
    var productDurations = {
      '402556P1': 3,
      '402556P3': 2.5,
      '402556P4': 2.5,
      '402556P5': 2.5,
      '402556P8': 2.5,
      '402556P7': 2,
      '402556P6': 2.5,
      '402556P9': 2.5,
      '402556P14': 4,
      '402556P10': 2.5,
      '402556P11': 3,
      '402556P12': 2.5,
      '402556P13': 3,
      '402556P16': 2.5,
      '402556P17': 2.5,
      '402556P19': 2,
      '402556P20': 3,
      '402556P21': 3,
      '402556P22': 2,
    };
    
    // Mapping of product codes to their respective Calendar IDs(GitHub note: Not shared for privacy reasons)
    var productCalendarMapping = {

    };
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var body = message.getPlainBody();
        var subject = message.getSubject();
        // Extract date, time, and product code
        var dateRegex = /(\d{1,2}\.\w{3} '\d{2}) @ (\d{1,2}:\d{2})/;
        var dateMatch = body.match(dateRegex) || subject.match(dateRegex);
        var bookingCodeRegex = /\(LOK-(.*?)\)/;
        var bookingCodeMatch = subject.match(bookingCodeRegex);
        var productRegex = /Product\s+([\w-]+)\s+-/;
        var productMatch = body.match(productRegex);
  
        var productNameRegex = /Product\s+\d{6}P\d+\s+-\s+(.*)/;
        var bookingChannelRegex = /^Booking channel\s+(.*)$/m;
        var customerRegex = /^Customer\s+(.*)$/m;
        var customerPhoneRegex = /^Customer phone\s+(\+\d+)$/m;
        var rateRegex = /^Rate\s+(.*)$/m;
        var productNameMatch = body.match(productNameRegex);
        var bookingChannelMatch = body.match(bookingChannelRegex);
        var customerMatch = body.match(customerRegex);
        var customerPhoneMatch = body.match(customerPhoneRegex);
        var rateMatch = body.match(rateRegex);
  
        if (dateMatch && productMatch && bookingCodeMatch) {
          var dateTimeString = dateMatch[1] + " " + dateMatch[2];
          var date = new Date(dateTimeString.replace(/'/g, "20"));
          var productCode = productMatch[1];
          var duration = productDurations[productCode] || 2; // Default duration if not listed
          var calendarId = productCalendarMapping[productCode];
          
          if (calendarId) {
            var calendar = CalendarApp.getCalendarById(calendarId);
            var startTime = date;
            var endTime = new Date(startTime.getTime() + (duration * 60 * 60 * 1000)); // Convert hours to milliseconds
            
            // Create or update calendar event
            var eventTitle = "Booking: " + (subject.includes("Cancelled booking:") ? "CANCELLED " : "") + dateTimeString + " " + bookingCodeMatch[0];
            var eventDescription = "Product: " + (productNameMatch ? productNameMatch[1] : "N/A") +
              "\nBooking channel: " + (bookingChannelMatch ? bookingChannelMatch[1] : "N/A") +
              "\nCustomer: " + (customerMatch ? customerMatch[1] : "N/A") +
              "\nCustomer phone: " + (customerPhoneMatch ? customerPhoneMatch[1] : "N/A") +
              "\nRate: " + (rateMatch ? rateMatch[1] : "N/A") +
              "\nDate: " + dateTimeString;
  
            var events = calendar.getEvents(startTime, endTime);
            var eventCheck = false;
            if (events.length > 0) {
              for(var k = 0; k < events.length;k++)
              {
                  var event = events[k];
                  if(event.getTitle() === eventTitle)
                  {
                    eventCheck = true;
                    break;
                  }
              }
              if(!eventCheck)
              {
                calendar.createEvent(eventTitle, startTime, endTime, {description: eventDescription});
              }
            } 
            else {
  
              // No events found, safe to create a new one
              calendar.createEvent(eventTitle, startTime, endTime, {description: eventDescription});
            }
            
           
          }
        }
      }
    }
  }