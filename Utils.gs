/**
 * Utils.gs - fungsi bantu lintas modul: perhitungan jarak GPS, reverse
 * geocoding, dan parsing tanggal/jam. Tidak ada state sheet di sini.
 */

// ============ JARAK GPS ============

function hitungJarakMeter_(lat1, lon1, lat2, lon2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}


// ============ REVERSE GEOCODING (Google Maps Geocoding API) ============

function reverseGeocode_(lat, lng) {
  var hasil = { kelurahan: '-', kecamatan: '-', kabupaten: '-' };
  var apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
  if (!apiKey) return hasil;

  try {
    var url = 'https://maps.googleapis.com/maps/api/geocode/json?latlng=' + lat + ',' + lng +
      '&key=' + apiKey + '&language=id';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (data.status === 'OK' && data.results && data.results.length > 0) {
      var comps = data.results[0].address_components;
      comps.forEach(function (c) {
        if (c.types.indexOf('administrative_area_level_4') !== -1 || c.types.indexOf('sublocality_level_1') !== -1) {
          hasil.kelurahan = c.long_name;
        }
        if (c.types.indexOf('administrative_area_level_3') !== -1) {
          hasil.kecamatan = c.long_name;
        }
        if (c.types.indexOf('administrative_area_level_2') !== -1) {
          hasil.kabupaten = c.long_name;
        }
      });
    } else {
      Logger.log('Geocoding status: ' + data.status);
    }
  } catch (e) {
    Logger.log('Geocoding error: ' + e.message);
  }
  return hasil;
}


// ============ PARSING TANGGAL & JAM ============

function parseTanggalDDMMYYYY_(str) {
  var bagian = str.split('/');
  if (bagian.length !== 3) return null;
  var d = parseInt(bagian[0], 10), m = parseInt(bagian[1], 10), y = parseInt(bagian[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m - 1, d);
}

// Tanggal ijin dikirim dari <input type="date"> (format yyyy-MM-dd), beda dari format tanggal absensi (dd/MM/yyyy)
function parseTanggalFleksibel_(str) {
  if (!str) return null;
  if (str.indexOf('-') !== -1) {
    var bag = str.split('-');
    if (bag.length === 3 && bag[0].length === 4) {
      var y = parseInt(bag[0], 10), m = parseInt(bag[1], 10), d = parseInt(bag[2], 10);
      if (!isNaN(y) && !isNaN(m) && !isNaN(d)) return new Date(y, m - 1, d);
    }
  }
  return parseTanggalDDMMYYYY_(str);
}

function formatJamSel_(cell) {
  if (!cell && cell !== 0) return '';
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(cell).trim().substring(0, 5);
}

function menitDariJam_(jamStr) {
  var bagian = String(jamStr || '').split(':');
  if (bagian.length < 2) return null;
  var jam = parseInt(bagian[0], 10);
  var menit = parseInt(bagian[1], 10);
  if (isNaN(jam) || isNaN(menit)) return null;
  return jam * 60 + menit;
}
