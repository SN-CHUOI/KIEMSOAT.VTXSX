// Worker nay nhan du lieu tu dashboard web, roi ghi de vao file data.js
// trong repo GitHub thay cho ban - de moi nguoi mo link deu thay du lieu
// moi nhat vinh vien, khong can chay file .bat tren may nao ca.
//
// Bao mat: GITHUB_TOKEN duoc luu rieng trong muc Settings -> Variables cua
// Worker (dang "Secret"), khong nam trong code nay, khong ai xem duoc.

const ALLOWED_ORIGIN = "https://sn-chuoi.github.io";
const OWNER = "SN-CHUOI";
const REPO = "KIEMSOATVATU.XSX";
const FILE_PATH = "data.js";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json; charset=utf-8"
  };
}

function b64EncodeUtf8(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  bytes.forEach(b => { binary += String.fromCharCode(b); });
  return btoa(binary);
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    const url = new URL(request.url);
    if (url.searchParams.get("debug") === "1") {
      const t = env.GITHUB_TOKEN || "";
      return new Response(JSON.stringify({
        tokenExists: !!env.GITHUB_TOKEN,
        tokenLength: t.length,
        tokenPrefix: t.slice(0, 11),
        tokenLast4: t.slice(-4)
      }), { status: 200, headers: corsHeaders() });
    }

    if (request.method !== "POST") {
      return new Response(JSON.stringify({ error: "Chi nhan POST" }), { status: 405, headers: corsHeaders() });
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return new Response(JSON.stringify({ error: "Du lieu gui len khong dung dinh dang JSON" }), { status: 400, headers: corsHeaders() });
    }

    const sources = body.sources;
    if (!Array.isArray(sources) || sources.length === 0) {
      return new Response(JSON.stringify({ error: "Danh sach nguon du lieu rong hoac khong hop le" }), { status: 400, headers: corsHeaders() });
    }

    const apiBase = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${FILE_PATH}`;
    const ghHeaders = {
      "Authorization": `Bearer ${env.GITHUB_TOKEN}`,
      "User-Agent": "kiemsoatvatu-publish-worker",
      "Accept": "application/vnd.github+json"
    };

    // 1. Lay sha hien tai cua file (bat buoc phai co de GitHub cho ghi de)
    let sha;
    try {
      const getResp = await fetch(apiBase, { headers: ghHeaders });
      if (!getResp.ok) {
        const t = await getResp.text();
        return new Response(JSON.stringify({ error: "Khong lay duoc file hien tai tu GitHub: " + t }), { status: 500, headers: corsHeaders() });
      }
      const getData = await getResp.json();
      sha = getData.sha;
    } catch (e) {
      return new Response(JSON.stringify({ error: "Loi ket noi GitHub (buoc doc file): " + e.message }), { status: 500, headers: corsHeaders() });
    }

    // 2. Dung noi dung data.js moi
    const now = new Date();
    const pad = n => String(n).padStart(2, "0");
    const genTimestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    const jsContent =
      "// File du lieu duoc cap nhat qua web - KHONG chinh sua tay\n" +
      `// Cap nhat luc: ${genTimestamp}\n` +
      `// So nguon: ${sources.length}\n` +
      `const DASHBOARD_DATA_SOURCES = ${JSON.stringify(sources)};\n`;

    // 3. Ghi de len GitHub
    try {
      const putResp = await fetch(apiBase, {
        method: "PUT",
        headers: { ...ghHeaders, "Content-Type": "application/json" },
        body: JSON.stringify({
          message: `Cap nhat du lieu tu web: ${body.sourceFileName || "khong ro ten file"}`,
          content: b64EncodeUtf8(jsContent),
          sha: sha
        })
      });
      if (!putResp.ok) {
        const t = await putResp.text();
        return new Response(JSON.stringify({ error: "Khong ghi duoc len GitHub: " + t }), { status: 500, headers: corsHeaders() });
      }
    } catch (e) {
      return new Response(JSON.stringify({ error: "Loi ket noi GitHub (buoc ghi file): " + e.message }), { status: 500, headers: corsHeaders() });
    }

    return new Response(JSON.stringify({ success: true }), { status: 200, headers: corsHeaders() });
  }
};
