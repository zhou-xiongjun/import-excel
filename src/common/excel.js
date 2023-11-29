import { saveAs } from "file-saver";
import * as EXCEL from "exceljs";

/**
 * 导出excel
 */
export async function exportExcel(
  dataArr, // 数据
  excelColumns, // 表头
  cellsStyle = {}, // 单元格样式
  filename = "表格"
) {
  /**
 * 
 * 
***********示例***********
let data = [
  { name: "张三", age: 18 },
  { name: "李四", age: 20 },
];
let excelColumns = [
  {
    header: "姓名",
    key: "name",
    style: {},
    width: 25,
  },
  {
    header: "年龄",
    key: "age",
    style: {},
    width: 25,
  },
];
let cellsStyle = { merge: ["A10:A13"] };
exportExcel(data, excelColumns, cellsStyle);

 * 
 */

  const workbook = new EXCEL.Workbook();
  const worksheet = workbook.addWorksheet("sheet1");
  worksheet.columns = excelColumns;
  dataArr.forEach((element) => {
    worksheet.addRow(element);
  });

  let { merge, cellStyleArr } = cellsStyle;

  if (merge && merge.length) {
    merge.forEach((item) => {
      worksheet.mergeCells(item);
    });
  }

  if (cellStyleArr && cellStyleArr.length) {
    cellStyleArr.forEach((item) => {
      let { cell, style } = item;
      cell.length &&
        cell.forEach((cellItem) => {
          for (const key in style) {
            worksheet.getCell(cellItem)[key] = style[key];
          }
        });
    });
  }

  workbook.xlsx.writeBuffer().then((data) => {
    saveAs(new Blob([data]), `${filename}.xlsx`);
  });
}

/**
 * 导入excel
 */
export async function importExcel(callback) {
  const workbook = new EXCEL.Workbook();
  const inputObj = document.createElement("input");
  document.getElementById("file")?.remove();
  inputObj.setAttribute("id", "file");
  inputObj.setAttribute("type", "file");
  inputObj.setAttribute("name", "file");
  inputObj.setAttribute("style", "display:none");
  inputObj.setAttribute("accept", ".xlsx, .xls");

  inputObj.addEventListener("change", (evt) => {
    const file = evt.target.files[0];
    let reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onloadend = function (e) {
      let data = new Uint8Array(e.target.result);
      workbook.xlsx.load(data).then(() => {
        if (workbook._themes.theme1) {
          workbook.eachSheet((worksheet) => {
            let excelData = [];
            worksheet.model.rows.forEach((row) => {
              excelData.push(row.cells.map((i) => i.value));
            });
            callback(excelData);
          });
        } else {
          console.log("文件格式错误");
        }
      });
    };
  });
  document.body.appendChild(inputObj);
  inputObj.click();
}

export function formatExcel(
  data,
  headKeyRowIdx,
  excelColumns,
  headLabelList = []
) {
  let headKeyObj = {};
  let headLabelArr = headLabelList ? headLabelList : data[headKeyRowIdx - 1];
  let headKeyIdxObj = {};
  let arr = [];

  excelColumns.forEach((item) => {
    headKeyObj[item.header] = item.key;
  });

  headLabelArr.forEach((item) => {
    if (headKeyObj[item]) {
      const idx = headLabelArr.indexOf(item);
      headKeyIdxObj[headKeyObj[item]] = idx;
    }
  });
  data.slice(headKeyRowIdx).forEach((item) => {
    let rowItem = {};
    for (const key in headKeyIdxObj) {
      const idx = headKeyIdxObj[key];
      rowItem[key] = item[idx];
    }
    arr.push(rowItem);
  });
  return arr;
}
