
const CONFIG = {
  SHEET_NAME: 'Rank History',
  QUERY: 'oprah winfrey',
  TARGET_IMAGE_URL: 'https://www.amacad.org/sites/default/files/person/headshots/oprah.jpg',
  TARGET_SOURCE_URL: 'https://www.amacad.org/person/oprah-winfrey',
  LOCATION_LABEL: 'United States — no city specified',
  GOOGLE_DOMAIN: 'google.com',
  COUNTRY: 'us',
  LANGUAGE: 'en',
  NO_CACHE: true,
  TRIGGER_HOUR: 7,
  TIMEZONE: 'America/New_York'
};

const HEADERS = [
  'Date and Time',
  'Query',
  'Device',
  'Location',
  'Image Carousel Present',
  'Target Image Found',
  'Image Position',
  'Number of Images',
  'Carousel Type',
  'Returned Image URL',
  'Returned Source Page',
  'Image Title',
  'Match Type',
  'Google SERP URL',
  'Notes'
];

function setupTrackingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);
  } else {
    const existingHeaders = sheet
      .getRange(1, 1, 1, HEADERS.length)
      .getValues()[0];

    if (existingHeaders.join('|') !== HEADERS.join('|')) {
      throw new Error(
        'Existing headers do not match the required 15 columns.'
      );
    }
  }

  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setWrap(true);

  sheet.setFrozenRows(1);

  sheet.getRange('A:A')
    .setNumberFormat('m/d/yyyy h:mm:ss AM/PM');

  sheet.autoResizeColumns(1, HEADERS.length);
}

function checkMobileAndDesktopRank() {
  checkImageRank('mobile');
  checkImageRank('desktop');
}

function checkMobileRank() {
  checkImageRank('mobile');
}

function checkDesktopRank() {
  checkImageRank('desktop');
}

function checkImageRank(device) {
  if (!['mobile', 'desktop'].includes(device)) {
    throw new Error('Invalid device: ' + device);
  }

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('SERPAPI_KEY');

  if (!apiKey) {
    throw new Error(
      'Missing SERPAPI_KEY in Script Properties.'
    );
  }

  const params = {
    engine: 'google',
    q: CONFIG.QUERY,
    start: 0,
    device: device,
    google_domain: CONFIG.GOOGLE_DOMAIN,
    gl: CONFIG.COUNTRY,
    hl: CONFIG.LANGUAGE,
    no_cache: String(CONFIG.NO_CACHE),
    output: 'json',
    api_key: apiKey
  };

  const queryString = Object.keys(params)
    .map(function(key) {
      return (
        encodeURIComponent(key) + '=' +
        encodeURIComponent(params[key])
      );
    })
    .join('&');

  const requestUrl =
    'https://serpapi.com/search.json?' + queryString;

  const response = UrlFetchApp.fetch(requestUrl, {
    method: 'get',
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      'SerpApi HTTP ' + status + ': ' +
      response.getContentText().slice(0, 500)
    );
  }

  const data = JSON.parse(response.getContentText());

  if (data.error) {
    throw new Error('SerpApi error: ' + data.error);
  }

  saveTrackingResult(data, device);
}

function getImageCollections(data) {
  const collections = [];

  if (
    Array.isArray(data.inline_images) &&
    data.inline_images.length > 0
  ) {
    collections.push({
      type: 'Inline image pack',
      images: data.inline_images
    });
  }

  if (
    data.knowledge_graph &&
    Array.isArray(data.knowledge_graph.header_images) &&
    data.knowledge_graph.header_images.length > 0
  ) {
    collections.push({
      type: 'Knowledge panel image carousel',
      images: data.knowledge_graph.header_images
    });
  }

  return collections;
}

function normalizeImageUrl(url) {
  if (!url) return '';

  try {
    const parsed = new URL(String(url).trim());

    return (
      parsed.hostname.replace(/^www\./, '').toLowerCase() +
      parsed.pathname.replace(/\/$/, '').toLowerCase()
    );
  } catch (error) {
    return String(url).trim().toLowerCase();
  }
}

function normalizeSourceUrl(url) {
  if (!url) return '';

  try {
    let parsed = new URL(String(url).trim());

    if (/google\./i.test(parsed.hostname)) {
      const redirect =
        parsed.searchParams.get('url') ||
        parsed.searchParams.get('q') ||
        parsed.searchParams.get('imgrefurl');

      if (redirect && /^https?:\/\//i.test(redirect)) {
        parsed = new URL(redirect);
      }
    }

    return (
      parsed.hostname.replace(/^www\./, '').toLowerCase() +
      parsed.pathname
        .replace(/\/amp\/?$/i, '')
        .replace(/\/$/, '')
        .toLowerCase()
    );
  } catch (error) {
    return String(url).trim().toLowerCase();
  }
}

function getFilename(url) {
  if (!url) return '';

  try {
    return decodeURIComponent(
      new URL(url).pathname.split('/').pop()
    ).toLowerCase();
  } catch (error) {
    return '';
  }
}

function getImageUrls(image) {
  return [
    image.original,
    image.image,
    image.thumbnail,
    image.serpapi_thumbnail,
    image.image_url
  ].filter(function(value) {
    return typeof value === 'string' && value.length > 0;
  });
}

function getReturnedImageUrl(image) {
  return getImageUrls(image)[0] || '';
}

function getSourceUrls(image) {
  const candidates = [
    image.link,
    image.source_url,
    image.page_url,
    image.original_link
  ];

  if (typeof image.source === 'string') {
    candidates.push(image.source);
  } else if (image.source) {
    candidates.push(
      image.source.link,
      image.source.url
    );
  }

  return candidates.filter(function(value) {
    return typeof value === 'string' &&
      /^https?:\/\//i.test(value);
  });
}

function getReturnedSourceUrl(image) {
  return getSourceUrls(image)[0] || '';
}

function getImagePosition(image, index) {
  return typeof image.position === 'number'
    ? image.position
    : index + 1;
}

function findTargetImage(images) {
  const targetUrl = normalizeImageUrl(
    CONFIG.TARGET_IMAGE_URL
  );

  const targetFilename = getFilename(
    CONFIG.TARGET_IMAGE_URL
  );

  const targetSource = normalizeSourceUrl(
    CONFIG.TARGET_SOURCE_URL
  );

  for (let i = 0; i < images.length; i++) {
    const image = images[i];

    for (const url of getImageUrls(image)) {
      if (normalizeImageUrl(url) === targetUrl) {
        return {
          image: image,
          position: getImagePosition(image, i),
          matchType: 'Exact image URL',
          confirmed: true
        };
      }

      if (
        targetFilename &&
        getFilename(url) === targetFilename
      ) {
        return {
          image: image,
          position: getImagePosition(image, i),
          matchType: 'Image filename',
          confirmed: true
        };
      }
    }
  }

  const sourceMatches = [];

  images.forEach(function(image, index) {
    const matchesSource = getSourceUrls(image)
      .some(function(url) {
        return normalizeSourceUrl(url) === targetSource;
      });

    if (matchesSource) {
      sourceMatches.push({
        image: image,
        position: getImagePosition(image, index)
      });
    }
  });

  if (sourceMatches.length === 1) {
    return {
      image: sourceMatches[0].image,
      position: sourceMatches[0].position,
      matchType: 'Source page URL fallback',
      confirmed: false,
      ambiguous: false
    };
  }

  if (sourceMatches.length > 1) {
    return {
      image: null,
      position: '',
      matchType: 'Multiple images from source page',
      confirmed: false,
      ambiguous: true
    };
  }

  return null;
}

function saveTrackingResult(data, device) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Run setupTrackingSheet() first.'
    );
  }

  const collections = getImageCollections(data);

  let selectedCollection = null;
  let selectedMatch = null;

  for (const collection of collections) {
    const match = findTargetImage(collection.images);

    if (match && match.confirmed) {
      selectedCollection = collection;
      selectedMatch = match;
      break;
    }
  }

  if (!selectedCollection) {
    for (const collection of collections) {
      const match = findTargetImage(collection.images);

      if (match) {
        selectedCollection = collection;
        selectedMatch = match;
        break;
      }
    }
  }

  if (!selectedCollection && collections.length) {
    selectedCollection = collections[0];
  }

  const images = selectedCollection
    ? selectedCollection.images
    : [];

  const confirmed = Boolean(
    selectedMatch && selectedMatch.confirmed
  );

  const sourceOnly = Boolean(
    selectedMatch &&
    !selectedMatch.confirmed &&
    !selectedMatch.ambiguous
  );

  const image = selectedMatch
    ? selectedMatch.image
    : null;

  let notes = '';

  if (!collections.length) {
    notes = 'No image carousel parsed in response.';
  } else if (confirmed) {
    notes = 'Exact image identified.';
  } else if (sourceOnly) {
    notes =
      'Source page matched, but exact image unverified. ' +
      'Position not recorded as confirmed.';
  } else if (
    selectedMatch &&
    selectedMatch.ambiguous
  ) {
    notes =
      'Multiple images from source page. ' +
      'Exact image position cannot be confirmed.';
  } else {
    notes =
      'Target image not identified in parsed carousel.';
  }

  const googleUrl =
    data.search_metadata &&
    data.search_metadata.google_url
      ? data.search_metadata.google_url
      : '';

  const row = [
    new Date(),
    CONFIG.QUERY,
    device,
    CONFIG.LOCATION_LABEL,
    collections.length ? 'Yes' : 'No',
    confirmed ? 'Yes' :
      sourceOnly ? 'Unverified' : 'No',
    confirmed ? selectedMatch.position : '',
    images.length,
    selectedCollection ? selectedCollection.type : '',
    image ? getReturnedImageUrl(image) : '',
    image ? getReturnedSourceUrl(image) : '',
    image ? (image.title || '') : '',
    selectedMatch ? selectedMatch.matchType : '',
    googleUrl,
    notes
  ];

  sheet.appendRow(row);
}

function createDailyTrigger() {
  deleteExistingRankTriggers();

  ScriptApp.newTrigger('checkMobileAndDesktopRank')
    .timeBased()
    .atHour(CONFIG.TRIGGER_HOUR)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
}

function deleteExistingRankTriggers() {
  const handlers = [
    'checkMobileAndDesktopRank',
    'checkMobileRank',
    'checkDesktopRank',
    'checkImagePackRank',
    'checkImagePackRankForDevice'
  ];

  ScriptApp.getProjectTriggers().forEach(
    function(trigger) {
      if (
        handlers.includes(
          trigger.getHandlerFunction()
        )
      ) {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  );
}
