function run() {
  const prop = PropertiesService.getScriptProperties().getProperties();
  const calendar = CalendarApp.getCalendarById(prop.calendar_id);

  // Push Gmail messages into an array
  const threads = GmailApp.search('yoyaku@expy.jp subject:{新幹線予約内容, 新幹線予約変更内容}');
  const messages = threads.reduceRight(function(a, b) {
    return a.concat(b.getMessages());
  }, []);
  
  messages.forEach(function(message) {
    if(message.getDate().getTime() < prop.last_execution) return;

    const body = message.getBody().split('変更前の予約内容');
    
    if(body.length === 2) {
      // Delete the old event
      const before = scrape(body[1], message.getDate());
      deleteEvent(calendar, before);
    }
    
    const bookings = body[0].split('■お申込内容')[0].split('【帰り】');    
    bookings.forEach(function(booking) {
      const reservation = scrape(booking, message.getDate());
      
      // Create an event
      addEvent(calendar, reservation)
    });
  });
  PropertiesService.getScriptProperties().setProperty('last_execution', Date.now());
}

function addEvent(calendar, reservation) {
  Logger.log('add')
  Logger.log(reservation)
  if (reservation.reserved) {
    calendar.createEvent(
      reservation.eventName,
      reservation.departure,
      reservation.arrival,
      {description: reservation.description}
    );
  } else {
    calendar.createAllDayEvent(
      reservation.eventName,
      reservation.departure,
      {description: reservation.description}
    );
  }
  Logger.log('success')
}

function getEvent(calendar, reservation) {
  if (reservation.reserved) {
    return calendar.getEvents(reservation.departure, reservation.arrival).filter(function(e) {
      return e.getTitle().indexOf('新幹線：') !== -1;
    })[0];
  } else {
    return calendar.getEventsForDay(reservation.departure).filter(function(e) {
      return e.getTitle().indexOf('新幹線：') !== -1;
    })[0];
  }
}

function deleteEvent(calendar, reservation) {
  Logger.log('delete')
  Logger.log(reservation)
  const event = getEvent(calendar, reservation)
  event.deleteEvent();
  Logger.log('success')
}

function scrape(text, messageDate) {
  const reserved = text.indexOf('自由席') === -1
  if (reserved) {
    return scrapeReserved(text, messageDate)
  } else {
    return scrapeNonReserved(text, messageDate)
  }
}

function scrapeReserved(text, messageDate) {
  var year = messageDate.getFullYear()
  const month = messageDate.getMonth()

  const date = text.match(/乗車日　(\d{1,2})月(\d{1,2})日/);
  
  if (month === 11 && date[1] === '1') {
    year += 1
  }
  
  const seat = text.match(/\d{1,2}号車\d{1,2}番[A-E]{1,5}席/);
  const seatNumber = seat[0]
  
  const train = text.match(/(.+)\((\d{1,2}):(\d{1,2})\)→(.+号)→(.+)\((\d{1,2}):(\d{1,2})\)/);
  const from = train[1]
  const to = train[5]
  const trainNumber = train[4]
    
  const departure = new Date(year, +date[1] - 1, +date[2], +train[2], +train[3])
  const arrival = new Date(year, +date[1] - 1, +date[2], +train[6], +train[7])
  const description = trainNumber + ' ' + seatNumber
      
  const eventName = '新幹線：' + from + '→' + to
  
  return {
    reserved: true,
    eventName: eventName,
    description: description,
    departure: departure,
    arrival: arrival
  };
}

function scrapeNonReserved(text, messageDate) {
  var year = messageDate.getFullYear()
  const month = messageDate.getMonth()

  const date = text.match(/乗車日　(\d{1,2})月(\d{1,2})日/);
  
  if (month === 11.0 && date[1] === '1') {
    year += 1
  }
  
  const seatNumber = '自由席'
  const description = seatNumber
  const departure = new Date(year, +date[1] - 1, +date[2], +0, +0)
  const arrival = new Date(year, +date[1] - 1, +date[2], +0, +0)
  
  const route = text.match(/(.+)→(.+)/);
  const from = route[1]
  const to = route[2]

  const eventName = '新幹線：' + from + '→' + to
  

  
  return {
    reserved: false,
    eventName: eventName,
    description: description,
    departure: departure,
    arrival: arrival
  };
}
