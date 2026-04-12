/**
 * Argus 수소·암모니아 PDF → OCR → 한글 번역 자동화 스크립트 v3
 *
 * 기능:
 *  1. Gmail에서 Argus PDF 첨부파일 수신 감지
 *  2. Google Drive PDF 폴더에 저장
 *  3. Drive OCR로 영문 텍스트 추출 → OCR 폴더에 저장
 *  4. LanguageApp으로 한글 번역 → 번역본 폴더에 저장
 *
 * 설정값 (아래 4개 폴더 ID만 본인 것으로 교체하세요)
 */

// ── 폴더 ID 설정 ─────────────────────────────────────────────
const PDF_FOLDER_ID   = '1oMyl4hTVN8chOw5MPLQKYWdMonStRLxo'; // 원문 PDF 저장 폴더
const OCR_FOLDER_ID   = '1R3F2gqKA4m4lKi7dA-E-f_vtXjqDLDug'; // OCR 영문 결과 폴더
const TRANS_FOLDER_ID = '1Gw18D61S2DFNG1MVnvvBTIf9Yr7hLjFH';   // 번역본 폴더
const PROCESSED_LABEL = 'Argus-processed';                     // 처리 완료 Gmail 라벨

// ── 메인 실행 함수 (트리거 연결) ─────────────────────────────
function processArgusEmails() {
  const label    = getOrCreateLabel(PROCESSED_LABEL);
  const threads  = GmailApp.search('from:notifications@argusmedia.com has:attachment -label:' + PROCESSED_LABEL, 0, 10);

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      msg.getAttachments().forEach(att => {
        const name = att.getName();
        if (!name.toLowerCase().endsWith('.pdf')) return;

        try {
          // 1) PDF 저장
          const pdfBlob   = att.copyBlob().setName(name);
          const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);

          // 월별 하위 폴더 자동 생성 (예: 2026-04)
          const monthKey    = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
          const monthFolder = getOrCreateSubfolder(pdfFolder, monthKey);
          const pdfFile     = monthFolder.createFile(pdfBlob);

          // 2) OCR → 영문 Google Doc
          const ocrDocId = ocrPdf(pdfFile.getId(), name.replace('.pdf', ''), OCR_FOLDER_ID);

          // 3) 한글 번역본 생성
          if (ocrDocId && TRANS_FOLDER_ID !== 'YOUR_TRANSLATION_FOLDER_ID') {
            translateDocToKorean(ocrDocId, name.replace('.pdf', ''), TRANS_FOLDER_ID);
          }

          Logger.log('완료: ' + name);
        } catch (e) {
          Logger.log('오류 [' + name + ']: ' + e.message);
        }
      });
    });

    // 처리 완료 라벨 추가
    thread.addLabel(label);
  });
}

// ── OCR 함수 ─────────────────────────────────────────────────
function ocrPdf(pdfFileId, baseName, targetFolderId) {
  const token   = ScriptApp.getOAuthToken();
  const mimeUrl = 'https://www.googleapis.com/drive/v3/files/' + pdfFileId + '?fields=mimeType';

  // Drive API v3로 OCR 변환 요청
  const uploadUrl = 'https://www.googleapis.com/drive/v3/files/copy';
  const payload   = JSON.stringify({
    name    : baseName + '_OCR',
    parents : [targetFolderId],
    mimeType: 'application/vnd.google-apps.document'
  });

  const resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + pdfFileId + '/copy',
    {
      method     : 'POST',
      contentType: 'application/json',
      headers    : { Authorization: 'Bearer ' + token },
      payload    : payload,
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(resp.getContentText());
  if (result.error) {
    Logger.log('OCR 오류: ' + JSON.stringify(result.error));
    return null;
  }
  Logger.log('OCR 완료: ' + result.id);
  return result.id;
}

// ── 한글 번역 함수 ───────────────────────────────────────────
/**
 * OCR Google Doc의 텍스트를 읽어 한글로 번역 후
 * 새 Google Doc으로 TRANS_FOLDER_ID에 저장
 */
function translateDocToKorean(ocrDocId, baseName, transFolderId) {
  const srcDoc  = DocumentApp.openById(ocrDocId);
  const body    = srcDoc.getBody();
  const srcText = body.getText();

  if (!srcText || srcText.trim().length === 0) {
    Logger.log('번역 건너뜀 — 빈 문서: ' + ocrDocId);
    return;
  }

  // Google Apps Script 내장 번역 함수 (영어 → 한국어)
  // 텍스트가 길면 5000자 단위로 분할 번역
  const CHUNK = 4500;
  let translated = '';

  if (srcText.length <= CHUNK) {
    translated = LanguageApp.translate(srcText, 'en', 'ko');
  } else {
    const chunks = [];
    for (let i = 0; i < srcText.length; i += CHUNK) {
      chunks.push(srcText.substring(i, i + CHUNK));
    }
    // 청크별 번역 (API 과부하 방지 위해 0.5초 간격)
    chunks.forEach((chunk, idx) => {
      if (idx > 0) Utilities.sleep(500);
      translated += LanguageApp.translate(chunk, 'en', 'ko') + '\n';
    });
  }

  // 번역본 문서 생성
  const transFolder = DriveApp.getFolderById(transFolderId);
  const dateStr     = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd');
  const newDocName  = baseName + '_번역본';

  const newDoc  = DocumentApp.create(newDocName);
  const newBody = newDoc.getBody();

  // 제목 + 원문 정보 헤더 추가
  newBody.appendParagraph('[Argus Media 한글 번역본]')
         .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  newBody.appendParagraph('원본: ' + baseName + '  |  번역일: ' + dateStr)
         .setFontSize(10);
  newBody.appendHorizontalRule();
  newBody.appendParagraph(translated);

  newDoc.saveAndClose();

  // 번역본 폴더로 이동
  const newFile = DriveApp.getFileById(newDoc.getId());
  transFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile); // 루트에서 제거

  Logger.log('번역본 생성 완료: ' + newDocName + ' (' + newDoc.getId() + ')');
  return newDoc.getId();
}

// ── 유틸리티 함수 ────────────────────────────────────────────
function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getOrCreateSubfolder(parentFolder, name) {
  const iter = parentFolder.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parentFolder.createFolder(name);
}

// ── 기존 OCR 파일 소급 번역 함수 (수동 1회 실행) ─────────────
/**
 * OCR 폴더에 있는 파일 중 번역본 폴더에 대응 파일이 없는 것만 골라
 * 한글 번역본을 생성합니다.
 *
 * 사용법: Apps Script 편집기에서 이 함수를 선택 후 ▶ 실행
 */
function translateExistingOcrDocs() {
  const ocrFolder   = DriveApp.getFolderById(OCR_FOLDER_ID);
  const transFolder = DriveApp.getFolderById(TRANS_FOLDER_ID);

  // 번역본 폴더에 이미 있는 파일명 목록 수집 (중복 방지)
  const existingNames = {};
  const existIter = transFolder.getFiles();
  while (existIter.hasNext()) {
    const f = existIter.next();
    existingNames[f.getName()] = true;
  }

  // OCR 폴더의 모든 Google Doc 순회
  const ocrIter = ocrFolder.getFiles();
  let count = 0;

  while (ocrIter.hasNext()) {
    const file = ocrIter.next();

    // Google Docs 형식만 처리
    if (file.getMimeType() !== 'application/vnd.google-apps.document') continue;

    const baseName    = file.getName().replace(/_OCR$/, '');
    const targetName  = baseName + '_번역본';

    // 이미 번역본이 있으면 건너뜀
    if (existingNames[targetName]) {
      Logger.log('건너뜀 (이미 번역됨): ' + targetName);
      continue;
    }

    Logger.log('번역 시작: ' + file.getName());

    try {
      translateDocToKorean(file.getId(), baseName, TRANS_FOLDER_ID);
      count++;
      // API 과부하 방지 — 파일 간 1초 대기
      Utilities.sleep(1000);
    } catch (e) {
      Logger.log('오류 [' + file.getName() + ']: ' + e.message);
    }
  }

  Logger.log('소급 번역 완료 — 총 ' + count + '건 생성');
}

// ── 트리거 설정 함수 (최초 1회 실행) ─────────────────────────
function setupTrigger() {
  // 기존 트리거 삭제 후 재설정
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processArgusEmails')
           .timeBased()
           .everyHours(1)          // 1시간마다 확인 (필요 시 everyMinutes(30) 등으로 변경)
           .create();
  Logger.log('트리거 설정 완료');
}
