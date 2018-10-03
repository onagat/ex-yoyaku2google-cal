function run() {
  const prop = PropertiesService.getScriptProperties().getProperties();
  const calendar = CalendarApp.getCalendarById(prop.calendar_id);

  // Push Gmail messages into an array
  const threads = GmailApp.search('yoyaku@expy.jp subject:新幹線予約内容');
  const messages = threads.reduceRight(function(a, b) {
    return a.concat(b.getMessages());
  }, []);

  messages.forEach(function(message) {
    if(message.getDate().getTime() < prop.last_execution) return;

    const body = message.getBody().split('変更前の予約内容');
    if(body[0].indexOf('自由席を予約しました') !== -1) return;

    const reservation = scrape(body[0]);

    if(body.length === 2) {
      // Delete the old event
      const before = scrape(body[1]);
      const event = calendar.getEvents(before.departure, before.arrival).filter(function(e) {
        return e.getTitle().indexOf('🚅') !== -1;
      })[0];
      event.deleteEvent();
    }

    // Create an event
    calendar.createEvent(
      '🚅' + reservation.train,
      reservation.departure,
      reservation.arrival,
      {location: reservation.seat}
    );
  });

  PropertiesService.getScriptProperties().setProperty('last_execution', Date.now());
}

function scrape(text) {
  const year = new Date().getFullYear();

  const date = text.match(/乗車日　(\d{1,2})月(\d{1,2})日/);
  const time = text.match(/.+\((\d{1,2}):(\d{1,2})\)→(.+号)→.+\((\d{1,2}):(\d{1,2})\)/);
  const seat = text.match(/\d{1,2}号車\d{1,2}番[A-E]席/);

  return {
    train: time[3],
    seat: seat[0],
    departure: new Date(year, +date[1] - 1, +date[2], +time[1], +time[2]),
    arrival: new Date(year, +date[1] - 1, +date[2], +time[4], +time[5]),
  };
}
