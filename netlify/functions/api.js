const { google } = require("googleapis");

const ADMIN_PW = "9600";

function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
}

async function getSheets() {
  const auth = getAuth();
  return google.sheets({ version: "v4", auth });
}

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

async function readSheet(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: sheetName,
  });
  return res.data.values || [];
}

async function appendRow(sheets, sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueInputOption: "USER_ENTERED",
    resource: { values: [values] },
  });
}

async function updateCell(sheets, sheetName, row, col, value) {
  const colLetter = String.fromCharCode(64 + col);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!${colLetter}${row}`,
    valueInputOption: "USER_ENTERED",
    resource: { values: [[value]] },
  });
}

async function updateRow(sheets, sheetName, row, values) {
  const endCol = String.fromCharCode(64 + values.length);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetName}!A${row}:${endCol}${row}`,
    valueInputOption: "USER_ENTERED",
    resource: { values: [values] },
  });
}

async function deleteRow(sheets, sheetName, rowIndex) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = meta.data.sheets.find(
    (s) => s.properties.title === sheetName
  );
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: sheet.properties.sheetId,
              dimension: "ROWS",
              startIndex: rowIndex - 1,
              endIndex: rowIndex,
            },
          },
        },
      ],
    },
  });
}

async function safeTranslate(text, targetLang) {
  if (!text || !targetLang) return text;
  if (targetLang === "ko" && /[가-힣]/.test(text.toString())) return text;
  try {
    const placeholder = " [n] ";
    const prepared = text.toString().replace(/\n/g, placeholder);
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${encodeURIComponent(targetLang)}&dt=t&q=${encodeURIComponent(prepared)}`;
    const res = await fetch(url);
    const json = await res.json();
    const translated = json[0].map((part) => part[0]).join("");
    return translated.replace(/ \[n\] /gi, "\n").replace(/\[n\]/gi, "\n");
  } catch {
    return text;
  }
}

function formatDate(dateVal) {
  const d = new Date(dateVal);
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getMonth() + 1)}.${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// --- 핸들러 함수들 ---

async function handleGetInitData({ grade, group }) {
  const sheets = await getSheets();
  const data = await readSheet(sheets, "Settings");
  const gNum = grade.toString().replace(/[^0-9]/g, "");
  const grNum = group.toString().replace(/[^0-9]/g, "");
  const filtered = data.slice(1)
    .filter((row) => {
      const rg = (row[1] || "").toString().replace(/[^0-9]/g, "");
      const rgr = (row[2] || "").toString().replace(/[^0-9]/g, "");
      return rg === gNum && rgr === grNum && (row[4] || "").toString().toUpperCase() === "ON";
    })
    .map((row) => ({ name: row[0].toString() }));
  return { classList: filtered };
}

async function handleGetPosts({ classTitle, grade, group, targetLang }) {
  const sheets = await getSheets();
  const [data, setRows] = await Promise.all([
    readSheet(sheets, "Data"),
    readSheet(sheets, "Settings"),
  ]);

  let ex = "", desc = "";
  for (let j = 1; j < setRows.length; j++) {
    if (setRows[j][0] === classTitle && setRows[j][1] === grade && setRows[j][2] === group) {
      desc = setRows[j][3] || "";
      ex = setRows[j][5] || "";
      break;
    }
  }

  const result = {
    title: await safeTranslate(classTitle, targetLang),
    desc: await safeTranslate(desc, targetLang),
    example: await safeTranslate(ex, targetLang),
    posts: [],
  };

  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === classTitle && data[i][2] === grade && data[i][3] === group) {
      let subj = data[i][6] || "";
      let msg = data[i][7] || "";
      if (targetLang === "zh") {
        subj = data[i][11] || (await safeTranslate(data[i][6], "zh"));
        msg = data[i][12] || (await safeTranslate(data[i][7], "zh"));
      } else if (targetLang === "ru") {
        subj = data[i][13] || (await safeTranslate(data[i][6], "ru"));
        msg = data[i][14] || (await safeTranslate(data[i][7], "ru"));
      } else if (targetLang === "ko") {
        subj = data[i][15] || (await safeTranslate(data[i][6], "ko"));
        msg = data[i][16] || (await safeTranslate(data[i][7], "ko"));
      }
      result.posts.push({
        id: data[i][0].toString(),
        num: data[i][4],
        name: data[i][5],
        subject: subj,
        message: msg,
        imgs: data[i][9] ? data[i][9].split("|") : [],
        likes: data[i][10] || 0,
        time: formatDate(data[i][0]),
        origSub: data[i][6] || "",
        origMsg: data[i][7] || "",
        origImgs: data[i][9] || "",
      });
    }
  }
  return result;
}

async function handleProcessForm(formObject) {
  const sheets = await getSheets();
  const [sub_zh, msg_zh, sub_ru, msg_ru, sub_ko, msg_ko] = await Promise.all([
    safeTranslate(formObject.subject, "zh"),
    safeTranslate(formObject.message, "zh"),
    safeTranslate(formObject.subject, "ru"),
    safeTranslate(formObject.message, "ru"),
    safeTranslate(formObject.subject, "ko"),
    safeTranslate(formObject.message, "ko"),
  ]);

  if (formObject.mode === "edit" || formObject.mode === "delete") {
    const data = await readSheet(sheets, "Data");
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === formObject.postId) {
        if (data[i][8].toString() === formObject.password || formObject.password === ADMIN_PW) {
          rowIdx = i + 1;
          break;
        }
      }
    }
    if (rowIdx < 0) throw new Error("비밀번호 불일치");
    if (formObject.mode === "edit") {
      await updateRow(sheets, "Data", rowIdx, [
        data[rowIdx - 1][0], data[rowIdx - 1][1], data[rowIdx - 1][2], data[rowIdx - 1][3],
        data[rowIdx - 1][4], data[rowIdx - 1][5],
        formObject.subject, formObject.message,
        data[rowIdx - 1][8], formObject.combinedImageData,
        data[rowIdx - 1][10],
        sub_zh, msg_zh, sub_ru, msg_ru, sub_ko, msg_ko,
      ]);
    } else {
      await deleteRow(sheets, "Data", rowIdx);
    }
  } else {
    const now = new Date().toISOString();
    await appendRow(sheets, "Data", [
      now, formObject.classTitle, formObject.grade, formObject.group,
      formObject.num, formObject.name, formObject.subject, formObject.message,
      formObject.password, formObject.combinedImageData, 0,
      sub_zh, msg_zh, sub_ru, msg_ru, sub_ko, msg_ko,
    ]);
  }
  return handleGetPosts({
    classTitle: formObject.classTitle,
    grade: formObject.grade,
    group: formObject.group,
    targetLang: formObject.targetLang,
  });
}

async function handleToggleLike({ postId, isCancel }) {
  const sheets = await getSheets();
  const data = await readSheet(sheets, "Data");
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === postId) {
      const cur = parseInt(data[i][10] || 0);
      const newVal = isCancel ? Math.max(0, cur - 1) : cur + 1;
      await updateCell(sheets, "Data", i + 1, 11, newVal);
      return newVal;
    }
  }
  return 0;
}

async function handleAddClass({ adminPw, classInfo }) {
  if (adminPw !== ADMIN_PW) throw new Error("비밀번호 불일치");
  const sheets = await getSheets();
  if (classInfo.row && classInfo.row !== "") {
    const data = await readSheet(sheets, "Settings");
    const rowNum = Number(classInfo.row);
    const existingStatus = (data[rowNum - 1] || [])[4] || "ON";
    await updateRow(sheets, "Settings", rowNum, [
      classInfo.name, classInfo.grade, classInfo.group,
      classInfo.desc, existingStatus, classInfo.ex,
    ]);
  } else {
    const targets = classInfo.targetClasses;
    if (!targets || targets.length === 0) throw new Error("대상 학급을 선택해주세요.");
    for (const t of targets) {
      await appendRow(sheets, "Settings", [
        classInfo.name, t.grade, t.group, classInfo.desc, "ON", classInfo.ex,
      ]);
    }
  }
  return true;
}

async function handleUpdateClassStatus({ pw, row, newStatus }) {
  if (pw !== ADMIN_PW) throw new Error("권한이 없습니다.");
  const sheets = await getSheets();
  await updateCell(sheets, "Settings", Number(row), 5, newStatus);
  return true;
}

async function handleGetSettingsList({ pw }) {
  if (pw !== ADMIN_PW) return [];
  const sheets = await getSheets();
  const data = await readSheet(sheets, "Settings");
  return data.slice(1).map((r, i) => ({
    row: i + 2,
    name: r[0] || "",
    grade: r[1] || "",
    group: r[2] || "",
    desc: r[3] || "",
    status: r[4] || "",
    ex: r[5] || "",
  }));
}

async function handleDeleteClass({ pw, row }) {
  if (pw !== ADMIN_PW) throw new Error("권한이 없습니다.");
  const sheets = await getSheets();
  await deleteRow(sheets, "Settings", Number(row));
  return true;
}

// --- 메인 핸들러 ---

exports.handler = async (event) => {
  const headers = {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers, body: JSON.stringify({ error: "Method Not Allowed" }) };
  }

  let body;
  try {
    body = JSON.parse(event.body || "{}");
  } catch {
    return { statusCode: 400, headers, body: JSON.stringify({ error: "Invalid JSON" }) };
  }

  const { fn, ...params } = body;

  try {
    let result;
    switch (fn) {
      case "getInitData":     result = await handleGetInitData(params); break;
      case "getPosts":        result = await handleGetPosts(params); break;
      case "processForm":     result = await handleProcessForm(params); break;
      case "toggleLike":      result = await handleToggleLike(params); break;
      case "addClass":        result = await handleAddClass(params); break;
      case "updateClassStatus": result = await handleUpdateClassStatus(params); break;
      case "getSettingsList": result = await handleGetSettingsList(params); break;
      case "deleteClass":     result = await handleDeleteClass(params); break;
      default:
        return { statusCode: 400, headers, body: JSON.stringify({ error: `Unknown function: ${fn}` }) };
    }
    return { statusCode: 200, headers, body: JSON.stringify({ ok: true, result }) };
  } catch (err) {
    return {
      statusCode: 400,
      headers,
      body: JSON.stringify({ ok: false, error: err.message }),
    };
  }
};
