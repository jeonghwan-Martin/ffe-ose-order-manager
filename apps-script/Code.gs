/**
 * FF&E/OS&E 발주관리 웹앱의 데이터 저장소 역할을 하는 Apps Script 백엔드.
 * PropertiesService(스크립트 속성)를 간단한 키-값 데이터베이스로 사용합니다.
 *
 * 배포 방법은 이 폴더의 README.md를 참고하세요.
 */

function doGet(e) {
  const action = e.parameter.action;
  const props = PropertiesService.getScriptProperties();

  if (action === "listProjects") {
    const list = props.getProperty("project-index") || "[]";
    return jsonResponse({ list: JSON.parse(list) });
  }

  if (action === "getProject") {
    const id = e.parameter.id;
    const value = props.getProperty("project-data:" + id) || "";
    return jsonResponse({ value });
  }

  if (action === "getCatalog") {
    const value = props.getProperty("item-catalog") || "";
    return jsonResponse({ value });
  }

  return jsonResponse({ error: "알 수 없는 action: " + action });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const props = PropertiesService.getScriptProperties();

  if (body.action === "saveProject") {
    props.setProperty("project-data:" + body.id, body.value);
    return jsonResponse({ ok: true });
  }

  if (body.action === "saveIndex") {
    props.setProperty("project-index", body.value);
    return jsonResponse({ ok: true });
  }

  if (body.action === "saveCatalog") {
    props.setProperty("item-catalog", body.value);
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: "알 수 없는 action: " + body.action });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
