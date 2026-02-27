function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.deployURL = ScriptApp.getService().getUrl();
  template.formHTML = getFormHTML('');
  const htmlOutput = template.evaluate();
  return htmlOutput;
}

function doPost(e) {
  let shainName = '';
  const shainCode = e.parameter.shaincode;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d'); // 今日の日付を取得

  if (!shainCode || shainCode.trim() === '') {
    const alert = '社員コードはありません';
    const template = HtmlService.createTemplateFromFile('index');
    template.deployURL = ScriptApp.getService().getUrl();
    template.formHTML = getFormHTML(alert);
    return template.evaluate();
  }

  if (e.parameter.createSanitary) {
    const syainCDsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    const textObject = syainCDsheet.createTextFinder(shainCode).matchEntireCell(true);
    const results = textObject.findAll();

    if (results.length > 0) {
      const rowIndex = results[0].getRow();
      shainName = syainCDsheet.getRange(rowIndex, 2).getValue();
      
      // 日付のシートを取得または作成
      let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(today);
      if (!dataSheet) {
        const emptySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('空'); // 「空」シートを参考にする
        dataSheet = emptySheet.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(today);
        dataSheet.showSheet(); // コピーしたシートを表示させる
      }


      const registeredRow = dataSheet.createTextFinder(shainCode).matchEntireCell(true).findNext();

      let employeeData = [];
      if (registeredRow) {
        const registeredRowIndex = registeredRow.getRow();
        employeeData = dataSheet.getRange(registeredRowIndex, 4, 1, dataSheet.getLastColumn()).getValues()[0];
      }

      const template = HtmlService.createTemplateFromFile('createSanitary');
      template.deployURL = ScriptApp.getService().getUrl();
      template.shainCode = shainCode;
      template.shainName = shainName;
      template.createSanitary = createSanitary(employeeData);
      const htmlOutput = template.evaluate();
      return htmlOutput;

    } else {
      const alert = '社員コードはありません';
      const template = HtmlService.createTemplateFromFile('index');
      template.deployURL = ScriptApp.getService().getUrl();
      template.formHTML = getFormHTML(alert);
      return template.evaluate();
    }
  }

  if (e.parameter.submit) {
    const shainCode = e.parameter.shaincode;
    const shainName = e.parameter.shainName;
    
    // 今日の日付のシートを取得または作成
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d');
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(today) || 
                      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('空').copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(today);
    
    const timestamp = new Date();
    let employeeData = [timestamp, shainCode, shainName];

    const item1 = e.parameter[`item1`] === '✔' ? '✔' : '';
    const item2 = e.parameter[`item2`] === '✔' ? '✔' : '';
    const item3 = e.parameter[`item3`] === '✔' ? '✔' : '';
    const item4 = e.parameter[`item4`] === '✔' ? '✔' : '';
    const item5 = e.parameter[`item5`] === '✔' ? '✔' : '';
    const item6 = e.parameter[`item6`] === '✔' ? '✔' : '';
    const item7 = e.parameter[`item7`] === '✔' ? '✔' : '';
    const item8 = e.parameter[`item8`] === '✔' ? '✔' : '';
    const item9 = e.parameter[`item9`] === '✔' ? '✔' : '';
    const item10 = e.parameter[`item10`] === '✔' ? '✔' : '';
    const item11 = e.parameter[`item11`] === '✔' ? '✔' : '';
    const item12 = e.parameter[`item12`] === '✔' ? '✔' : '';
    const item13 = e.parameter[`item13`] === '✔' ? '✔' : '';
    const item14 = e.parameter[`item14`] === '✔' ? '✔' : '';
    const item15 = e.parameter[`item15`] === '✔' ? '✔' : '';
    const item16 = e.parameter[`item16`] === '✔' ? '✔' : '';
    const item17 = e.parameter[`item17`] === '✔' ? '✔' : '';
    const item18 = e.parameter[`item18`] === '✔' ? '✔' : '';
    const item19 = e.parameter[`item19`] === '✔' ? '✔' : '';
    const item20 = e.parameter[`item20`] === '✔' ? '✔' : '';
    const item21 = e.parameter[`item21`] === '✔' ? '✔' : '';
    const item22 = e.parameter[`item22`] === '✔' ? '✔' : '';
    const item23 = e.parameter[`item23`] === '✔' ? '✔' : '';
    const item24 = e.parameter[`item24`] === '✔' ? '✔' : '';
    const item25 = e.parameter[`item25`] === '✔' ? '✔' : '';
    const item26 = e.parameter[`item26`] === '✔' ? '✔' : '';
    const text1 = e.parameter[`text1`] || '';
    const text2 = e.parameter[`text2`] || '';
    employeeData.push(item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15, item16, item17, item18, item19, item20, item21, item22, item23, item24, item25, item26, text1, text2);

    const registeredRow = dataSheet.createTextFinder(shainCode).matchEntireCell(true).findNext();
    if (registeredRow) {
      const registeredRowIndex = registeredRow.getRow();
      dataSheet.getRange(registeredRowIndex, 1, 1, employeeData.length).setValues([employeeData]);
    } else {
      dataSheet.appendRow(employeeData);
    }

    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(3, 1, lastRow - 1, dataSheet.getLastColumn());
    const data = dataRange.getValues();

    const absenceSheet = SpreadsheetApp.openById('1EpPcQdZG-dyDqsBrUyQDjhzS7tnnBc6josAVTrZ5OVk').getSheetByName('登録データ');

    if (absenceSheet.getLastRow() > 2) {
      const clearRange = absenceSheet.getRange(3, 1, absenceSheet.getLastRow() - 1, absenceSheet.getLastColumn());
      clearRange.clearContent();
    }

    absenceSheet.getRange(3, 1, data.length, data[0].length).setValues(data);

    const sortRange = absenceSheet.getRange(3, 1, absenceSheet.getLastRow() - 1, absenceSheet.getLastColumn());
    sortRange.sort({column: 3, ascending: true});

    const template = HtmlService.createTemplateFromFile('index');
    template.deployURL = ScriptApp.getService().getUrl();
    template.formHTML = getFormHTML('登録しました。');
    return template.evaluate();
  }

  if (e.parameter.modoru) {
    const template = HtmlService.createTemplateFromFile('index');
    template.deployURL = ScriptApp.getService().getUrl();
    template.formHTML = getFormHTML('');
    const htmlOutput = template.evaluate();
    return htmlOutput;
  }
}

function createSanitary(employeeData) {
  let employeeDataIndex = 0;
  const hasEmployeeData = Array.isArray(employeeData);
  let calendarHtml = ``;

  // 日付のシートを参照
  const item1 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item2 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item3 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item4 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item5 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item6 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item7 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item8 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item9 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item10 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item11 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item12 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item13 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item14 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item15 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item16 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item17 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item18 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item19 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item20 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item21 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item22 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item23 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item24 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item25 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const item26 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] === '✔' : false;
  const text1 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] : '';
  const text2 = hasEmployeeData && employeeDataIndex < employeeData.length ? employeeData[employeeDataIndex++] : '';
  
  calendarHtml += `<div style="display: block;"><h5>●日常確認（該当で『✔』 ）Kiểm tra hàng ngày (đánh dấu ✔ nếu có liên quan)</h5></div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item1" value="✔" ${item1 ? 'checked' : ''}> 手洗い・ローラー掛を実施した（Đã thực hiện rửa tay và dùng con lăn.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item2" value="✔" ${item2 ? 'checked' : ''}> 作業服・靴等の状態（Tình trạng trang phục làm việc và giày dép.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item3" value="✔" ${item3 ? 'checked' : ''}> 作業服・靴等着用状況（Tình trạng mặc trang phục làm việc và giày dép.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item4" value="✔" ${item4 ? 'checked' : ''}> ひげ・過度な化粧等（Râu và trang điểm quá mức.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item5" value="✔" ${item5 ? 'checked' : ''}> 爪の長さ（Chiều dài của móng tay.）</div>`;
  calendarHtml += `<div style="display: block;"><br></div>`;
  calendarHtml += `<div style="display: block;"><h5>●健康状態（該当で『✔』 ）Tình trạng sức khỏe (đánh dấu ✔ nếu có liên quan)</h5></div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item6" value="✔" ${item6 ? 'checked' : ''}> 貧血、高・低血圧がある（Có thiếu máu, huyết áp cao hoặc thấp.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item7" value="✔" ${item7 ? 'checked' : ''}> 発熱(37.0)以上ある（Có sốt (trên 37.0 độ).）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item8" value="✔" ${item8 ? 'checked' : ''}> 下痢・吐き気がある（Có dấu hiệu tiêu chảy hoặc buồn nôn.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item9" value="✔" ${item9 ? 'checked' : ''}> 手指怪我がある（Có thương tích ở tay.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item10" value="✔" ${item10 ? 'checked' : ''}> 歯科治療 治療中（Đang trong quá trình điều trị nha khoa.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item11" value="✔" ${item11 ? 'checked' : ''}> 歯科治療 異常がない（Không có bất thường trong điều trị nha khoa.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item12" value="✔" ${item12 ? 'checked' : ''}> 黄疸がある（Có dấu hiệu vàng da.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item13" value="✔" ${item13 ? 'checked' : ''}> 目、鼻分泌物がある（Có chất nhầy ở mắt và mũi.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item14" value="✔" ${item14 ? 'checked' : ''}> その他風邪の症状がある（Có triệu chứng cảm cúm khác.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item15" value="✔" ${item15 ? 'checked' : ''}> 腰痛がある（Có đau lưng.）</div>`;
  calendarHtml += `<div style="display: block;"><br></div>`;
  calendarHtml += `<div style="display: block;"><h5>●持ち込み確認（該当で『✔』 ）Kiểm tra đồ mang theo (đánh dấu ✔ nếu có liên quan)</h5></div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item16" value="✔" ${item16 ? 'checked' : ''}> メガネ（Kính mắt.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item17" value="✔" ${item17 ? 'checked' : ''}> ヘアピン（Kẹp tóc.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item18" value="✔" ${item18 ? 'checked' : ''}> ヘアゴム（dây buộc tóc.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item19" value="✔" ${item19 ? 'checked' : ''}> 義歯等（Răng giả.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item20" value="✔" ${item20 ? 'checked' : ''}> 医療器具等（Dụng cụ y tế.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item21" value="✔" ${item21 ? 'checked' : ''}> 会社指定絆創膏（Băng vết thương theo quy định của công ty.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item22" value="✔" ${item22 ? 'checked' : ''}> フォークリフト免許証（Giấy phép lái xe nâng.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item23" value="✔" ${item23 ? 'checked' : ''}> 薬等（Thuốc.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item24" value="✔" ${item24 ? 'checked' : ''}> コンタクト 使用者（Người sử dụng kính áp tròng.）</div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item25" value="✔" ${item25 ? 'checked' : ''}> コンタクト 紛失した者（Người đã mất kính áp tròng.）</div>`;
  calendarHtml += `<div style="display: block;"><br></div>`;
  calendarHtml += `<div style="display: block;"><h5>●指導者記入欄（Ô ghi thông tin của người hướng dẫn）</h5></div>`;
  calendarHtml += `<div style="display: block;"><input type="checkbox" name="item26" value="✔" ${item26 ? 'checked' : ''}> 指導有（Có chỉ đạo.）</div>`;
  calendarHtml += `<div style="display: block;">指導実施者（Người thực hiện chỉ đạo.）<br><input type="text" name="text1" placeholder="テキスト入力" value="${text1}" style="width: 80%; max-width: 100%;"></div>`;
  calendarHtml += `<div style="display: block;"><br></div>`;
  calendarHtml += `<div style="display: block;">補記（Ghi chú thêm.）<br><input type="text" name="text2" placeholder="テキスト入力" value="${text2}" style="width: 80%; max-width: 100%;"></div>`;
  
  return calendarHtml; // HTMLを戻り値として返す
}

function getFormHTML(alert = '') {
  let html = `
  <div class="mb-3">
    <label for="shaincode" class="form-label">社員コード</label>
    <input type="text" class="form-control" id="shaincode" name="shaincode">
  </div>
  <p class="text-danger">${alert}</p>
  `;
  return html;
}
