<script lang="ts">
  import { Spreadsheet } from 'tauri-plugin-spreadsheet-api';

  let spreadsheet = new Spreadsheet('E:\\zhidal\\Desktop\\test.xlsx', 'Sheet1');

  function read() {
    console.log('打开文件: ', spreadsheet.path);
    spreadsheet
      .read()
      .then((res) => {
        console.log('读取文件结果: ', res);
      })
      .catch((error) => {
        console.error(error);
      });
  }

  function create() {
    console.log('创建文件: ', spreadsheet.path);
    spreadsheet
      .create()
      .then((res) => {
        console.log('创建文件结果: ', res);
      })
      .catch((error) => {
        console.error(error);
      });
  }

  function write() {
    console.log('写入文件: ', spreadsheet.path);
    spreadsheet
      .write()
      .then((res) => {
        console.log('写入文件结果: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function getSheetRange() {
    console.log('获取 sheet 最大行列范围');
    spreadsheet
      .getSheetRange()
      .then((res) => {
        console.log('获取 sheet 最大宽高结果: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function getSheetRow() {
    console.log('获取 sheet 最大行');
    spreadsheet
      .getSheetRow()
      .then((res) => {
        console.log('获取 sheet 最大行结果: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function getSheetColumn() {
    console.log('获取 sheet 最大列');
    spreadsheet
      .getSheetColumn()
      .then((res) => {
        console.log('获取 sheet 最大列结果: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function newSheet() {
    const sheetName = 'Sheet2';
    console.log('创建新 sheet: ', sheetName);
    spreadsheet
      .newSheet(sheetName)
      .then((res) => {
        console.log('创建新 sheet: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function copytSheet() {
    const sheetName = 'CopySheet';
    console.log('复制 sheet: ', sheetName);
    spreadsheet
      .copySheet(sheetName)
      .then((res) => {
        console.log('复制 sheet: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function setValue() {
    const value = 'test';
    const local = [1, 1];
    console.log('设置值: ', local, value);
    spreadsheet
      .setValue(local, value)
      .then((res) => {
        console.log('设置值: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function getValue() {
    const local = [1, 1];
    console.log('获取值: ', local);
    spreadsheet
      .getValue(local)
      .then((res) => {
        console.log('获取值: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function removeRow() {
    const start = 1;
    const length = 2;
    console.log('移除行: ', start, length);
    spreadsheet
      .removeRow(start, length)
      .then((res) => {
        console.log('移除行: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function close() {
    console.log('关闭 xlsx');
    spreadsheet
      .close()
      .then((res) => {
        console.log('关闭 xlsx: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function closeAll() {
    console.log('关闭所有 xlsx 文件');
    Spreadsheet.closeAll()
      .then((res) => {
        console.log('关闭所有 xlsx 文件结果', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }

  function list() {
    console.log('列出所有 xlsx 文件: ');
    Spreadsheet.list()
      .then((res) => {
        console.log('列出所有 xlsx 文件结果: ', res);
      })
      .catch((err) => {
        console.error(err);
      });
  }
</script>

<div>
  <button on:click={read}>read</button>
  <button on:click={create}>create</button>
  <button on:click={write}>write</button>
  <button on:click={getSheetRange}>getSheetRange</button>
  <button on:click={getSheetColumn}>getSheetColumn</button>
  <button on:click={getSheetRow}>getSheetRow</button>
  <button on:click={copytSheet}>copytSheet</button>
  <button on:click={getValue}>getValue</button>
  <button on:click={setValue}>setValue</button>
  <button on:click={removeRow}>removeRow</button>
  <button on:click={close}>close</button>
  <button on:click={closeAll}>closeAll</button>
  <button on:click={list}>list</button>
</div>
