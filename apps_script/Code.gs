function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project EDGE')
    .addItem('Refresh Odds', 'refreshOdds')
    .addItem('Test Setup', 'testSetup')
    .addToUi();
}

function testSetup() {
  SpreadsheetApp.getUi().alert('Project EDGE is connected and working.');
}

function refreshOdds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const liveOdds = ss.getSheetByName('Live Odds');

  const apiKey = settings.getRange('B4').getValue() || settings.getRange('B2').getValue();
  const regions = settings.getRange('B5').getValue() || 'us,ca';
  const markets = settings.getRange('B6').getValue() || 'h2h';
  const oddsFormat = settings.getRange('B7').getValue() || 'american';

  const sports = [
    settings.getRange('B15').getValue(),
    settings.getRange('B16').getValue(),
    settings.getRange('B17').getValue(),
    settings.getRange('B18').getValue(),
    settings.getRange('B19').getValue()
  ].filter(String);

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Missing API key. Add it to Settings tab, cell B4 or B2.');
    return;
  }

  liveOdds.clearContents();

  const headers = [
    'Sport',
    'Game',
    'Start Time',
    'Bookmaker',
    'Market',
    'Outcome',
    'Odds',
    'Implied Probability',
    'Model Probability',
    'EV %'
  ];

  liveOdds.getRange(1, 1, 1, headers.length).setValues([headers]);

  let rows = [];

  sports.forEach(function(sportKey) {
    const url =
      'https://api.the-odds-api.com/v4/sports/' +
      encodeURIComponent(sportKey) +
      '/odds/?apiKey=' +
      encodeURIComponent(apiKey) +
      '&regions=' +
      encodeURIComponent(regions) +
      '&markets=' +
      encodeURIComponent(markets) +
      '&oddsFormat=' +
      encodeURIComponent(oddsFormat) +
      '&dateFormat=iso';

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();

    if (status !== 200) {
      rows.push([
        sportKey,
        'API error ' + status,
        '',
        '',
        '',
        response.getContentText().slice(0, 100),
        '',
        '',
        '',
        ''
      ]);
      return;
    }

    const events = JSON.parse(response.getContentText());

    events.forEach(function(event) {
      const game = event.away_team + ' @ ' + event.home_team;

      event.bookmakers.forEach(function(bookmaker) {
        bookmaker.markets.forEach(function(market) {
          market.outcomes.forEach(function(outcome) {
            const odds = Number(outcome.price);
            const implied = americanToImpliedProbability(odds);
            const modelProb = implied + 0.01;
            const ev = calculateEV(odds, modelProb);

            rows.push([
              sportKey,
              game,
              event.commence_time,
              bookmaker.key,
              market.key,
              outcome.name,
              odds,
              implied,
              modelProb,
              ev
            ]);
          });
        });
      });
    });
  });

  if (rows.length > 0) {
    liveOdds.getRange(2, 1, rows.length, headers.length).setValues(rows);
    liveOdds.getRange(2, 8, rows.length, 3).setNumberFormat('0.00%');
  }

  SpreadsheetApp.getUi().alert('Odds refresh complete. Rows loaded: ' + rows.length);
}

function americanToImpliedProbability(odds) {
  if (odds > 0) {
    return 100 / (odds + 100);
  }
  return Math.abs(odds) / (Math.abs(odds) + 100);
}

function calculateEV(americanOdds, modelProbability) {
  let profitPerDollar;

  if (americanOdds > 0) {
    profitPerDollar = americanOdds / 100;
  } else {
    profitPerDollar = 100 / Math.abs(americanOdds);
  }

  const loseProbability = 1 - modelProbability;
  return modelProbability * profitPerDollar - loseProbability;
}
