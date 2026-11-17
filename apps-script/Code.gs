/**
 * ORIGIN 2026 — Live Voting API (Google Apps Script)
 *
 * Cara pakai:
 * 1. Buka Google Spreadsheet > Extensions > Apps Script, lalu tempel seluruh isi file ini.
 * 2. Jalankan fungsi `setup` sekali (membuat sheet "Votes" beserta contoh datanya).
 * 3. Deploy > New deployment > Web app
 *      - Execute as     : Me
 *      - Who has access : Anyone
 * 4. Salin URL yang berakhiran /exec ke konstanta API_URL di vote.html.
 *
 * Struktur sheet "Votes" (urutan kolom bebas, yang dibaca adalah nama header-nya):
 *   Team | Votes | Image | Total Voters
 *
 *   - Team         : nama tim / project
 *   - Votes        : jumlah vote tim tersebut
 *   - Image        : (opsional) URL gambar, path relatif di website (mis. img/vote/tactile.png),
 *                    atau link share Google Drive
 *   - Total Voters : (opsional, cukup diisi di baris 2) jumlah seluruh pemilih.
 *                    Persentase = Votes / Total Voters.
 *                    Jika kosong, persentase = Votes / jumlah seluruh Votes.
 */

const SHEET_NAME = 'Votes';
const CACHE_KEY = 'vote-results';
const CACHE_SECONDS = 2; // meredam beban saat banyak layar membuka halaman bersamaan

function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get(CACHE_KEY);

  if (!json) {
    json = JSON.stringify(buildResults());
    cache.put(CACHE_KEY, json, CACHE_SECONDS);
  }

  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildResults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { ok: false, error: 'Sheet "' + SHEET_NAME + '" tidak ditemukan', results: [] };
  }

  const rows = sheet.getDataRange().getValues();
  const header = (rows.shift() || []).map(function (h) {
    return String(h).trim().toLowerCase();
  });

  const col = {
    team: header.indexOf('team'),
    votes: header.indexOf('votes'),
    image: header.indexOf('image'),
    total: header.indexOf('total voters'),
  };

  if (col.team < 0 || col.votes < 0) {
    return { ok: false, error: 'Header "Team" dan "Votes" wajib ada di baris 1', results: [] };
  }

  const teams = rows
    .map(function (row) {
      return {
        name: String(row[col.team]).trim(),
        votes: Math.max(0, Number(row[col.votes]) || 0),
        image: col.image >= 0 ? normalizeImageUrl(String(row[col.image]).trim()) : '',
      };
    })
    .filter(function (team) {
      return team.name !== '';
    });

  const sumVotes = teams.reduce(function (sum, team) {
    return sum + team.votes;
  }, 0);
  const totalVoters = col.total >= 0 && rows.length ? Number(rows[0][col.total]) || 0 : 0;
  const divisor = totalVoters > 0 ? totalVoters : sumVotes;

  teams.forEach(function (team) {
    const percent = divisor > 0 ? (team.votes / divisor) * 100 : 0;
    team.percent = Math.min(100, Math.round(percent));
  });

  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    totalVoters: divisor,
    results: teams,
  };
}

/** Link share Google Drive diubah menjadi URL gambar langsung. */
function normalizeImageUrl(url) {
  const match = url.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:.*&)?id=)([\w-]+)/);
  return match ? 'https://lh3.googleusercontent.com/d/' + match[1] : url;
}

/** Jalankan sekali untuk membuat sheet "Votes" beserta contoh datanya. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET_NAME)) {
    throw new Error('Sheet "' + SHEET_NAME + '" sudah ada — setup dibatalkan agar data tidak tertimpa.');
  }

  const sheet = ss.insertSheet(SHEET_NAME);
  const data = [
    ['Team', 'Votes', 'Image', 'Total Voters'],
    ['PROJECT TACTILE', 0, '', ''],
    ['TIGA HATI BERBAGI', 0, '', ''],
    ['BLOOMERS', 0, '', ''],
    ['BAWAH 25', 0, '', ''],
    ['KAPAW', 0, '', ''],
    ['HYDROPPY', 0, '', ''],
    ['SAHABATSEKOLAH', 0, '', ''],
  ];

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, data[0].length);
}
