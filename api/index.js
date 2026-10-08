import { createClient } from '@supabase/supabase-js';
import { google } from 'googleapis';
import stream from 'stream';

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
  },
  scopes: SCOPES,
});

const sheets = google.sheets({ version: 'v4', auth });
const drive = google.drive({ version: 'v3', auth });

const SPREADSHEET_ID = '16oz716kIVD3bA1OVnNZv0wrn7uVvm2Tw596W5i2gF78';
const BASE_FOLDER_ID = '1tAj_xlkQivJm8V--_Y3wzFVP72xm9V2h'; 

async function getSheetId(sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheet) throw new Error("Sheet tidak ditemukan. Pastikan nama Sheet di Excel sesuai.");
  return sheet.properties.sheetId;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Hanya POST yang diizinkan' });

  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    const { action, loginId, password } = payload;
    const aktorLogin = payload.aktorLogin || loginId;
    const targetSheet = payload.targetSheet || aktorLogin;

    if (action === 'login') {
      let { data: user } = await supabase.from('users')
        .select('*').or(`sheet_name.eq.${loginId},username.eq.${loginId}`).eq('password', password).single();

      if (!user) {
        try {
          await getSheetId(loginId);
          const { data: newUser, error: insertErr } = await supabase.from('users').insert([{
            sheet_name: loginId, username: loginId, password: password, role: 'User'
          }]).select().single();
          if (insertErr) throw insertErr;
          user = newUser;
        } catch (e) {
          return res.json({ error: "Username/Sheet tidak ditemukan atau Password salah." });
        }
      }
      return res.json({ ok: true, message: "Login berhasil", sheetName: user.sheet_name });
    }

    if (action === 'getBawahan') {
      const { data: me } = await supabase.from('users').select('*').eq('sheet_name', aktorLogin).single();
      const { data: allUsers } = await supabase.from('users').select('*');
      
      let bawahanList = [];
      const myRole = me?.role || 'User';
      const myUnit = me?.unit_kerja || '';
      const myJabatan = (me?.jabatan || '').toLowerCase();

      if (myRole === 'Super Admin' || myRole === 'Admin') {
        for (let u of allUsers) {
          if (u.sheet_name === aktorLogin) continue;
          let uJabatan = (u.jabatan || '').toLowerCase();
          
          if (myRole === 'Super Admin') {
            bawahanList.push({ sheet: u.sheet_name, name: u.username, role: u.role, unit: u.unit_kerja });
          } else if (myRole === 'Admin') {
            if (myJabatan.includes('sekretaris') && uJabatan.includes('kasubbag')) {
              bawahanList.push({ sheet: u.sheet_name, name: u.username, role: u.role, unit: u.unit_kerja });
            } else if (myJabatan.includes('kasubbag')) {
              let isSugiono = u.username.toLowerCase().includes('sugiono');
              let isHukumOrKeuangan = myUnit.toLowerCase().includes('hukum dan sdm') || myUnit.toLowerCase().includes('keuangan umum');
              
              if (u.unit_kerja === myUnit || (isHukumOrKeuangan && isSugiono)) {
                bawahanList.push({ sheet: u.sheet_name, name: u.username, role: u.role, unit: u.unit_kerja });
              }
            }
          }
        }
      }
      return res.json({ ok: true, bawahan: bawahanList, role: myRole });
    }

    if (action === 'initial') {
      const { data: user } = await supabase.from('users').select('*').eq('sheet_name', targetSheet).single();
      const response = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: SPREADSHEET_ID,
        ranges: [`${targetSheet}!C5:C9`, `${targetSheet}!A12:G1000`]
      });
      
      const pCells = response.data.valueRanges[0].values || [];
      const rCells = response.data.valueRanges[1].values || [];
      
      const profile = {
        nama: pCells[0]?.[0]?.replace(/^:\s*/, '') || '',
        nip: pCells[1]?.[0]?.replace(/^:\s*/, '') || '',
        jabatan: pCells[2]?.[0]?.replace(/^:\s*/, '') || '',
        unitKerja: pCells[3]?.[0]?.replace(/^:\s*/, '') || '',
        bulanLaporan: pCells[4]?.[0]?.replace(/^:\s*/, '') || 'Oktober 2026',
        username: user?.username || targetSheet,
        atasanTitle: user?.atasan_title || '',
        atasanName: user?.atasan_name || '',
        atasanNip: user?.atasan_nip || ''
      };

      const rows = rCells.map((r, i) => ({
        rowNumber: 12 + i, nomor: r[0] || '', tanggal: r[1] || '', pukul: r[2] || '',
        uraian: r[3] || '', jumlah: r[4] || '', link: r[5] || '', keterangan: r[6] || ''
      })).filter(r => r.nomor || r.tanggal || r.uraian); 

      return res.json({ profile, rows });
    }

    if (action === 'saveProfile') {
      const { profile } = payload;
      let updateData = {
        atasan_title: profile.atasanTitle, atasan_name: profile.atasanName,
        atasan_nip: profile.atasanNip, username: profile.username || targetSheet,
        unit_kerja: profile.unitKerja, jabatan: profile.jabatan
      };
      if (profile.password) updateData.password = profile.password;

      const { data: currUser } = await supabase.from('users').select('role').eq('sheet_name', targetSheet).single();
      if (currUser && currUser.role !== 'Super Admin') {
        const j = String(profile.jabatan).toLowerCase();
        updateData.role = (j.includes('sekretaris') || j.includes('kasubbag')) ? 'Admin' : 'User';
      }

      await supabase.from('users').update(updateData).eq('sheet_name', targetSheet);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${targetSheet}!C5:C9`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[`: ${profile.nama}`], [`: ${profile.nip}`], [`: ${profile.jabatan}`], [`: ${profile.unitKerja}`], [`: ${profile.bulanLaporan}`]] }
      });
      return res.json({ ok: true });
    }

    if (action === 'saveRow') {
      const sheetId = await getSheetId(targetSheet);
      let rowNumber = Number(payload.rowNumber);
      
      if (payload.insertAfter) {
        rowNumber = Number(payload.insertAfter) + 1;
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: [{ insertDimension: { range: { sheetId, dimension: "ROWS", startIndex: rowNumber - 1, endIndex: rowNumber }, inheritFromBefore: true } }]
          }
        });
      }

      let jamGabungan = payload.pukulAwal && payload.pukulAkhir ? `${payload.pukulAwal}-${payload.pukulAkhir} WIB` : payload.pukul;
      const rowValues = payload.isLanjutan 
        ? [['', '', '', payload.uraian || '', payload.jumlah || '', payload.link || '', payload.keterangan || '']]
        : [[payload.nomor || '', `'${payload.tanggalStr || ''}`, jamGabungan || '', payload.uraian || '', payload.jumlah || '', payload.link || '', payload.keterangan || '']];

      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID, range: `${targetSheet}!A${rowNumber}:G${rowNumber}`,
        valueInputOption: 'USER_ENTERED', requestBody: { values: rowValues }
      });

      const batchRequests = [];
      if (payload.isLanjutan) {
        for (let col = 0; col < 3; col++) {
          batchRequests.push({ mergeCells: { range: { sheetId, startRowIndex: rowNumber - 2, endRowIndex: rowNumber, startColumnIndex: col, endColumnIndex: col + 1 }, mergeType: "MERGE_ALL" } });
        }
      }
      batchRequests.push({ repeatCell: { range: { sheetId, startRowIndex: rowNumber - 1, endRowIndex: rowNumber, startColumnIndex: 0, endColumnIndex: 7 }, cell: { userEnteredFormat: { horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" } }, fields: "userEnteredFormat(horizontalAlignment,verticalAlignment)" } });
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: batchRequests } });
      
      return res.json({ ok: true });
    }

    if (action === 'deleteRow') {
      const sheetId = await getSheetId(targetSheet);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: Number(payload.rowNumber) - 1, endIndex: Number(payload.rowNumber) } } }] }
      });
      return res.json({ ok: true });
    }

    if (action === 'exportPdf') {
      const token = await auth.getAccessToken();
      const sheetId = await getSheetId(targetSheet);
      const exportUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?exportFormat=pdf&format=pdf&gid=${sheetId}&size=A4&portrait=true&fitw=true&gridlines=false`;
      
      const pdfRes = await fetch(exportUrl, { headers: { 'Authorization': `Bearer ${token}` } });
      const arrayBuffer = await pdfRes.arrayBuffer();
      const base64 = Buffer.from(arrayBuffer).toString('base64');
      return res.json({ dataUri: `data:application/pdf;base64,${base64}` });
    }

    if (action === 'simpanPdfKeDrive') {
      const folderName = payload.bulanLaporan || 'Tanpa Bulan';
      let folderId;
      
      const qFolder = `'${BASE_FOLDER_ID}' in parents and name = '${folderName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
      const searchRes = await drive.files.list({ q: qFolder, fields: 'files(id)' });
      
      if (searchRes.data.files.length > 0) {
        folderId = searchRes.data.files[0].id;
      } else {
        const createRes = await drive.files.create({ requestBody: { name: folderName, mimeType: 'application/vnd.google-apps.folder', parents: [BASE_FOLDER_ID] }, fields: 'id' });
        folderId = createRes.data.id;
      }

      const fileName = `LHK_${targetSheet}_${folderName}.pdf`;
      const qOldFile = `'${folderId}' in parents and name = '${fileName}' and trashed = false`;
      const oldFiles = await drive.files.list({ q: qOldFile });
      for (let f of oldFiles.data.files) {
        await drive.files.update({ fileId: f.id, requestBody: { trashed: true } });
      }

      const pdfBuffer = Buffer.from(payload.pdfBase64.split(',')[1], 'base64');
      const bufferStream = new stream.PassThrough();
      bufferStream.end(pdfBuffer);
      
      const uploadRes = await drive.files.create({ requestBody: { name: fileName, parents: [folderId] }, media: { mimeType: 'application/pdf', body: bufferStream }, fields: 'webViewLink' });
      return res.json({ ok: true, url: uploadRes.data.webViewLink });
    }

    return res.json({ error: "Aksi tidak dikenali." });

  } catch (err) {
    return res.status(500).json({ error: err.message || String(err) });
  }
}
