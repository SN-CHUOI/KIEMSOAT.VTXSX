// Ham server (chay tren Vercel) - giu GITHUB_TOKEN bi mat, nhan du lieu tu
// cong cu web roi ghi de vao file data.js trong repo GitHub thay cho nguoi dung.
// Khong ai (ke ca nguoi dung cong cu) can go token rieng nua.

const OWNER = 'SN-CHUOI';
const REPO = 'KIEMSOAT.VTXSX';
const FILE_PATH = 'data.js';
const ALLOWED_ORIGIN = 'https://sn-chuoi.github.io';

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN);
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Chi nhan POST' });
    return;
  }

  const token = process.env.GITHUB_TOKEN;
  if (!token) {
    res.status(500).json({ error: 'Server chua duoc cau hinh GITHUB_TOKEN (vao Vercel Settings > Environment Variables de them)' });
    return;
  }

  const sources = req.body && req.body.sources;
  if (!Array.isArray(sources) || sources.length === 0) {
    res.status(400).json({ error: 'Danh sach nguon du lieu rong hoac khong hop le' });
    return;
  }

  const apiBase = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${FILE_PATH}`;
  const ghHeaders = {
    Authorization: `Bearer ${token}`,
    'User-Agent': 'kiemsoatvatu-publish-fn',
    Accept: 'application/vnd.github+json'
  };

  try {
    const getResp = await fetch(apiBase, { headers: ghHeaders });
    if (!getResp.ok) {
      const t = await getResp.text();
      res.status(500).json({ error: 'Khong lay duoc file hien tai tu GitHub: ' + t });
      return;
    }
    const getData = await getResp.json();
    const sha = getData.sha;

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const ts = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    const jsContent =
      '// File du lieu duoc cap nhat qua cong cu web - KHONG chinh sua tay\n' +
      `// Cap nhat luc: ${ts}\n` +
      `// So nguon: ${sources.length}\n` +
      `const DASHBOARD_DATA_SOURCES = ${JSON.stringify(sources)};\n`;

    const contentB64 = Buffer.from(jsContent, 'utf-8').toString('base64');

    const putResp = await fetch(apiBase, {
      method: 'PUT',
      headers: { ...ghHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: 'Cap nhat du lieu qua cong cu web: ' + (req.body.sourceFileName || 'khong ro ten file'),
        content: contentB64,
        sha
      })
    });
    if (!putResp.ok) {
      const t = await putResp.text();
      res.status(500).json({ error: 'Khong ghi duoc len GitHub: ' + t });
      return;
    }

    res.status(200).json({ success: true });
  } catch (e) {
    res.status(500).json({ error: 'Loi ket noi GitHub: ' + e.message });
  }
};
