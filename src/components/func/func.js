// import { XLSX } from 'xlsx';
import { saveAs } from 'file-saver';
import XLSX from 'xlsx-js-style';
import ExcelJS from 'exceljs';

// Declare this so our linter knows that tableau is a global object
/* global tableau */


//array_move function from stackoverflow solution:
//https://stackoverflow.com/questions/5306680/move-an-array-element-from-one-array-position-to-another
function array_move(arr, old_index, new_index) {
  if (new_index >= arr.length) {
    var k = new_index - arr.length + 1;
    while (k--) {
      arr.push(undefined);
    }
  }
  arr.splice(new_index, 0, arr.splice(old_index, 1)[0]);
  return arr; // for testing
};

const saveSettings = () => new Promise((resolve, reject) => {
  console.log('[func.js] Saving settings');
  tableau.extensions.settings.set('metaVersion', 2);
  console.log('[func.js] Authoring mode', tableau.extensions.environment.mode);
  if (tableau.extensions.environment.mode === "authoring") {
    tableau.extensions.settings.saveAsync()
      .then(newSavedSettings => {
        //console.log('[func.js] newSavedSettings', newSavedSettings);
        resolve(newSavedSettings);
      }).catch(reject);
  } else {
    resolve();
  }

});

const setSettings = (type, value) => new Promise((resolve, reject) => {
  console.log('[func.js] Set settings', type, value);
  let settingKey = '';
  switch (type) {
    case 'sheets':
      settingKey = 'selectedSheets';
      break;
    case 'label':
      settingKey = 'buttonLabel';
      break;
    case 'style':
      settingKey = 'buttonStyle';
      break;
    case 'filename':
      settingKey = 'filename';
      break;
    case 'version':
      settingKey = 'metaVersion';
      break;
    default:
      settingKey = 'unknown';
  }
  tableau.extensions.settings.set(settingKey, JSON.stringify(value));
  resolve();
});

const getSheetColumns = (sheet, existingCols, modified) => new Promise((resolve, reject) => {
  sheet.getSummaryDataAsync({ ignoreSelection: true }).then((data) => {
    //console.log('[func.js] Sheet Summary Data', data);
    console.log('[func.js] getSheetColumns existingCols', JSON.stringify(existingCols));
    const columns = data.columns;
    let cols = [];
    const existingIdx = [];
    if (modified) {
      for (var j = 0; j < columns.length; j++) {
        //console.log(columns[j]);
        var col = {};
        col.index = columns[j].index;
        col.name = columns[j].fieldName;
        col.dataType = columns[j].dataType;
        col.isImage = columns[j].isImage;
        col.changeName = null;
        col.selected = false;
        cols.push(col);
      }
      for (var i = 0; i < existingCols.length; i++) {
        if (existingCols[i] && existingCols[i].hasOwnProperty("name")) {
          existingIdx.push(existingCols[i].name);
        }
      }
      console.log('[func.js] getSheetColumns existingIdx', existingIdx);
      let maxPos = existingIdx.length;
      cols = cols.map((col, idx) => {
        //console.log('[func.js] getSheetColumns Looking for col', col);
        const eIdx = existingIdx.indexOf(col.name);
        const ret = { ...col };
        if (eIdx > -1) {
          ret.selected = existingCols[eIdx].selected;
          ret.changeName = existingCols[eIdx].changeName;
          ret.isImage = existingCols[eIdx].isImage;
          ret.order = eIdx;
        } else {
          ret.order = maxPos;
          maxPos += 1;
        }
        return ret;
      });
    } else {
      for (var k = 0; k < columns.length; k++) {
        var newCol = {};
        newCol.index = columns[k].index;
        newCol.name = columns[k].fieldName;
        newCol.dataType = columns[k].dataType;
        newCol.isImage = columns[k].isImage;
        newCol.selected = true;
        newCol.order = k + 1;
        cols.push(newCol);
      }
    }
    cols = cols.sort((a, b) => (a.order > b.order) ? 1 : -1)
    resolve(cols);
  })
    .catch(error => {
      console.log('[func.js] Error with getSummaryDataAsync', sheet, error);
    });
});

const initializeMeta = () => new Promise((resolve, reject) => {
  console.log('[func.js] Initialise Meta');
  var promises = [];
  const worksheets = tableau.extensions.dashboardContent._dashboard.worksheets;
  //console.log('[func.js] Worksheets in dashboard', worksheets);
  var meta = worksheets.map(worksheet => {
    var sheet = worksheet;
    var item = {};
    item.sheetName = sheet.name;
    item.selected = false;
    item.changeName = null;
    item.customCols = false;
    promises.push(getSheetColumns(sheet, null, false));
    return item;
  });

  console.log(`[func.js] Found ${meta.length} sheets`, meta);

  Promise.all(promises).then((sheetArr) => {
    for (var i = 0; i < sheetArr.length; i++) {
      var sheetMeta = meta[i];
      sheetMeta.columns = sheetArr[i];
      meta[i] = sheetMeta;
      console.log(`[func.js] Added ${sheetArr[i].length} columns to ${sheetMeta.sheetName}`, meta);
    }
    //console.log(`[func.js] Meta initialised`, meta);
    resolve(meta);
  });
});

const revalidateMeta = (existing) => new Promise((resolve, reject) => {
  console.log('[func.js] Revalidate Meta');
  var promises = [];
  const worksheets = tableau.extensions.dashboardContent._dashboard.worksheets;
  //console.log('[func.js] Worksheets in dashboard', worksheets);
  var meta = worksheets.map(worksheet => {
    var sheet = worksheet;
    const sheetIdx = existing.findIndex((e) => {
      return e.sheetName === sheet.name;
    });
    if (sheetIdx > -1) {
      console.log(`[func.js] Existing sheet ${sheet.name} columns`, JSON.stringify(existing[sheetIdx].columns));
      promises.push(getSheetColumns(sheet, existing[sheetIdx].columns, true));
      existing[sheetIdx].existed = true;
      return existing[sheetIdx];
    } else {
      var item = {};
      item.sheetName = sheet.name;
      item.selected = false;
      item.changeName = null;
      item.customCols = false;
      item.existed = false;
      promises.push(getSheetColumns(sheet, null, false));
      return item;
    }
  });

  console.log(`[func.js] Found ${meta.length} sheets`, meta);

  Promise.all(promises).then((sheetArr) => {

    for (var i = 0; i < sheetArr.length; i++) {
      var sheetMeta = meta[i];
      sheetMeta.columns = sheetArr[i];
      meta[i] = sheetMeta;
      //console.log(`[func.js] Added ${sheetArr[i].length} columns to ${sheetMeta.sheetName}`, meta);
    }
    meta.forEach((sheet, idx) => {
      if (sheet && sheet.sheetName) {
        const eIdx = existing.findIndex((e) => {
          return e.sheetName === sheet.sheetName;
        });
        meta = array_move(meta, idx, eIdx);
      } else {
        console.log('[func.js] Sheet ordering issue. No sheet defined in idx', idx);
      }
    })
    console.log(`[func.js] Meta revalidated`, JSON.stringify(meta));
    resolve(meta);
  });
});

const exportToExcel = (meta, env, filename) => new Promise((resolve, reject) => {
  let xlsFile = "export.xlsx";
  if (filename && filename.length > 0) {
    xlsFile = filename + ".xlsx";
  }
  buildExcelBlob(meta).then(wb => {
    // add ignoreEC:false to prevent excel crashes during text to column
    var wopts = { bookType: 'xlsx', bookSST: false, type: 'array', ignoreEC: false };
    var wbout = XLSX.write(wb, wopts);
    saveAs(new Blob([wbout], { type: "application/octet-stream" }), xlsFile);
    resolve();
  });
});



// krisd: move excel creation to caller (to support extra export to methodss)
// callback receives a blob to save or transfer
const buildExcelBlob = (meta) => new Promise((resolve, reject) => {
  console.log("[func.js] Got Meta", meta);
  // func.saveSettings(meta, function(newSettings) {
  // console.log("Saved settings", newSettings);
  const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
  const wb = XLSX.utils.book_new();
  let totalSheets = 0;
  let sheetCount = 0;
  const sheetList = [];
  const columnList = [];
  const tabNames = [];
  for (let i = 0; i < meta.length; i++) {
    if (meta[i] && meta[i].selected) {
      let tabName = meta[i].changeName || meta[i].sheetName;
      tabName = tabName.replace(/[*?/\\[\]]/gi, '');
      sheetList.push(meta[i].sheetName);
      tabNames.push(tabName);
      columnList.push(meta[i].columns);
      totalSheets = totalSheets + 1;
    }
  }
  sheetList.map((metaSheet, idx) => {
    //console.log("[func.js] Finding sheet", metaSheet, worksheets);
    const sheet = worksheets.find(s => s.name === metaSheet);
    // eslint-disable-next-line
    sheet.getSummaryDataAsync({ ignoreSelection: true }).then((data) => {
      const columns = data.columns;
      const columnMeta = columnList[sheetCount];
      const headerOrder = [];
      columnMeta.map((colMeta, idx) => {
        if (colMeta && colMeta.selected) {
          headerOrder.push(colMeta.changeName || colMeta.name);
        }
        return colMeta;
      });
      columns.map((column, idx) => {
        //console.log("[func.js] Finding column", column.fieldName, columnMeta);
        const objCol = columnMeta.find(o => o.name === column.fieldName);
        if (objCol) {
          let col = { ...column, selected: objCol.selected }
          col.outputName = objCol.changeName || objCol.name;
          columns[idx] = col;
          return col;
        } else {
          return null;
        }
      });
      //console.log("[func.js] Running decodeRows", columns, data.data);
      decodeDataset(columns, data.data)
        .then((rows) => {
          //console.log("[func.js] decodeRows returned", rows);
          console.log("[func.js] Header Order", headerOrder);
          var ws = XLSX.utils.json_to_sheet(rows, { header: headerOrder });
          var sheetname = tabNames[sheetCount];
          sheetCount = sheetCount + 1;
          XLSX.utils.book_append_sheet(wb, ws, sheetname);
          if (sheetCount === totalSheets) {
            resolve(wb);
          }
        });
    });
    return sheet;
  });
});


// krisd: Remove recursion to work with larger data sets
// and translate cell data types
const decodeDataset = (columns, dataset) => new Promise((resolve, reject) => {
  let promises = [];
  for (let i = 0; i < dataset.length; i++) {
    promises.push(decodeRow(columns, dataset[i]));
  }
  Promise.all(promises).then((datasetArr) => {
    //console.log('[func.js] datasetArr', datasetArr);
    resolve(datasetArr);
  });

});

const decodeRow = (columns, row) => new Promise((resolve, reject) => {
  let meta = {};
  for (let j = 0; j < columns.length; j++) {
    if (columns[j].selected) {
      // krisd: let's assign the sheetjs type according to the summary data column type
      let dtype = undefined;
      let dval = undefined;
      // console.log('[func.js] Row', row[j]);
      if (row[j].value === '%null%' && row[j].nativeValue === null && row[j].formattedValue === 'Null') {
        dtype = 'z';
        dval = null;
      } else {
        switch (columns[j]._dataType) {
          case 'int':
          case 'float':
            dtype = 'n';
            dval = Number(row[j].value);  // let nums be raw w/o formatting
            if (isNaN(dval)) dval = row[j].formattedValue;  // protect in case issue
            break;
          case 'date':
          case 'date-time':
            dtype = 's';
            dval = row[j].formattedValue;
            break;
          case 'bool':
            dtype = 'b';
            dval = row[j].value;
            break;
          default:
            dtype = 's';
            dval = row[j].formattedValue;
        }
      }
      let o = { v: dval, t: dtype };
      meta[columns[j].outputName] = o;
    }
  }
  resolve(meta);
});



const downloadCrosstab = async (meta, filename) => {
  console.log('[func.js] Downloading Crosstab');
  let xlsFile = "crosstab_export.xlsx";
  if (filename && filename.length > 0) {
    xlsFile = filename + ".xlsx";
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await buildCrosstabExcelWorkbook(workbook, meta);
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), xlsFile);
  } catch (error) {
    console.error('Error generating Excel file:', error);
  }
};

const buildCrosstabExcelWorkbook = async (workbook, meta) => {
  console.log("[func.js] Got Meta for Crosstab", meta);
  const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;

  for (let i = 0; i < meta.length; i++) {
    if (meta[i] && meta[i].selected) {
      let tabName = meta[i].changeName || meta[i].sheetName;
      tabName = tabName.replace(/[*?/\\[\]]/gi, '');
      const sheet = worksheets.find(s => s.name === meta[i].sheetName);
      if (sheet) {
        await processSheet(workbook, sheet, tabName, meta[i]);
      }
    }
  }
};

const processSheet = async (workbook, sheet, tabName, sheetMeta) => {
  const worksheet = workbook.addWorksheet(tabName);
  const sheetData = await generateCrossTab(sheet);

  // ヘッダーの設定
  worksheet.columns = sheetData[0].map((header, index) => ({
    header,
    key: `col${index}`,
    width: 15
  }));

  // データの設定
  for (let i = 1; i < sheetData.length; i++) {
    const row = worksheet.addRow(sheetData[i]);
    row.height = 90; // 画像を表示するために行の高さを設定

    for (let j = 0; j < sheetData[i].length; j++) {
      const cell = row.getCell(j + 1);
      const columnMeta = sheetMeta.columns.find(col => col.name === sheetData[0][j]);

      console.log(cell);
      console.log(columnMeta);
      console.log(columnMeta?.isImage);
      if (columnMeta && columnMeta.isImage) {
        try {
          const imageUrl = sheetData[i][j];
          const response = await fetch(imageUrl);
          const blob = await response.blob();
          const arrayBuffer = await blob.arrayBuffer();

          const imageId = workbook.addImage({
            buffer: arrayBuffer,
            extension: 'jpeg',
          });

          worksheet.addImage(imageId, {
            tl: { col: j, row: i },
            ext: { width: 80, height: 80 },
            editAs: 'oneCell'
          });
        } catch (error) {
          console.error('Error fetching image:', error);
          cell.value = 'Image Load Error';
        }
      } else {
        cell.value = sheetData[i][j];
      }

      // スタイルの適用
      cell.font = {
        name: 'Meiryo UI',
        size: 9
      };
    }
  }

  // ヘッダー行のスタイル
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' }
  };
};

const buildCrosstabExcelBlob = (meta) => new Promise((resolve, reject) => {
  console.log("[func.js] Got Meta for Crosstab", meta);
  const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
  const wb = XLSX.utils.book_new();
  let totalSheets = 0;
  let sheetCount = 0;
  const sheetList = [];
  const tabNames = [];

  for (let i = 0; i < meta.length; i++) {
    if (meta[i] && meta[i].selected) {
      let tabName = meta[i].changeName || meta[i].sheetName;
      tabName = tabName.replace(/[*?/\\[\]]/gi, '');
      sheetList.push(meta[i].sheetName);
      tabNames.push(tabName);
      totalSheets = totalSheets + 1;
    }
  }

  const processSheet = async (metaSheet, idx) => {
    const sheet = worksheets.find(s => s.name === metaSheet);
    try {
      const sheetData = await generateCrossTab(sheet);
      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // 列幅の設定
      ws['!cols'] = sheetData[0].map(() => ({ wch: 15 }));

      // 各セルにフォントスタイルを適用
      for (let R = 0; R < sheetData.length; ++R) {
        for (let C = 0; C < sheetData[R].length; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[cellAddress]) ws[cellAddress] = {};
          ws[cellAddress].s = {
            font: {
              name: 'Meiryo UI',
              sz: 9
            }
          };
        }
      }

      const sheetname = tabNames[sheetCount];
      XLSX.utils.book_append_sheet(wb, ws, sheetname);
      sheetCount++;
      if (sheetCount === totalSheets) {
        resolve(wb);
      }
    } catch (error) {
      console.error('Error processing sheet:', error);
      reject(error);
    }
  };

  sheetList.forEach(processSheet);
});

async function generateCrossTab(worksheet) {
  const dataTable = await worksheet.getSummaryDataAsync();
  const visualSpec = await worksheet.getVisualSpecificationAsync();

  const rowFields = visualSpec.rowFields.map(f => f.name);
  const columnFields = visualSpec.columnFields.map(f => f.name);
  const measureField = visualSpec.marksSpecifications[0].encodings[0].field.name;

  const tempDataMap = new Map();

  dataTable.data.forEach(row => {
    const rowKey = rowFields.map(field => row[dataTable.columns.findIndex(col => col.fieldName === field)]._formattedValue).join('|');
    const colKey = columnFields.map(field => row[dataTable.columns.findIndex(col => col.fieldName === field)]._formattedValue).join('|');
    const value = Math.round(Number(row[dataTable.columns.findIndex(col => col.fieldName === measureField)].value) * 100) / 100;

    if (!tempDataMap.has(rowKey)) {
      tempDataMap.set(rowKey, new Map());
    }
    tempDataMap.get(rowKey).set(colKey, value);
  });

  const dataMap = new Map([...tempDataMap].reverse());

  for (let [key, value] of dataMap) {
    dataMap.set(key, new Map([...value].reverse()));
  }

  const uniqueColKeys = Array.from(new Set(dataTable.data.map(row =>
    columnFields.map(field => row[dataTable.columns.findIndex(col => col.fieldName === field)]._formattedValue).join('|')
  ))).reverse();

  let sheetData = [];

  // Create header rows
  for (let index = 0; index < columnFields.length; index++) {
    if (index === columnFields.length - 1) {
      const lastHeaderRow = rowFields.concat(uniqueColKeys.map(key => key.split('|')[index]));
      sheetData.push(lastHeaderRow);
    } else {
      const headerRow = new Array(rowFields.length).fill('').concat(uniqueColKeys.map(key => key.split('|')[index]));
      sheetData.push(headerRow);
    }
  }

  for (let [rowKey, rowData] of dataMap) {
    let row = rowKey.split('|');
    for (let colKey of uniqueColKeys) {
      row.push(rowData.get(colKey) || '');
    }
    sheetData.push(row);
  }

  console.log(sheetData)

  return sheetData;
}

export {
  initializeMeta,
  revalidateMeta,
  saveSettings,
  setSettings,
  exportToExcel,
  downloadCrosstab,
}
