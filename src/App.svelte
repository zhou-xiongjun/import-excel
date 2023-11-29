<script>
  import { exportExcel, importExcel, formatExcel } from "./common/excel";
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
  let importColumns = {
    name: "姓名",
    age: "年龄",
  };

  /**
   * 导入
   */
  const onImportFile = () => {
    let callback = (data) => {
      console.log(data);
      let Ddata = formatExcel(
        data,
        1,
        excelColumns,
        Object.values(importColumns)
      );
      console.log(Ddata);
    };
    importExcel(callback);
  };

  /**
   * 导出
   */
  const onExportFile = () => {
    let data = [
      { name: "张三", age: 18 },
      { name: "李四", age: 20 },
    ];
    let cellsStyle = { merge: ["A10:A13"] };
    exportExcel(data, excelColumns, cellsStyle);
  };

  /**
   * 模板下载
   */
  const onDownloadFile = () => {
    exportExcel([], excelColumns, {}, "模板");
  };
</script>

<main>
  <button on:click={onImportFile}>导入excel</button>
  <button on:click={onExportFile}>导出excel</button>
  <button on:click={onDownloadFile}>下载excel模板</button>
</main>
