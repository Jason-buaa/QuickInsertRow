/* global Office, Excel */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 插入按钮事件
    document.getElementById("insertBtn").onclick = () => {
      const entireRow = document.getElementById("toggleMode").checked;
      insertCells(entireRow);
    };
  }
});

/**
 * 插入空单元格/行（在当前单元格上方）
 * @param {boolean} entireRow - true = 插入整行；false = 仅当前列
 */
async function insertCells(entireRow) {
  const n = parseInt(document.getElementById("rowCount").value, 10);
  if (!Number.isInteger(n) || n <= 0) {
    return showStatus("❌ Please enter a valid positive number.");
  }

  try {
    await Excel.run(async (context) => {
      const active = context.workbook.getActiveCell();
      active.load("rowIndex, columnIndex, worksheet");
      await context.sync();

      const sheet = active.worksheet;

      if (entireRow) {
        // 插入整行：从当前行开始往下推
        const startRow = active.rowIndex;
        const endRow = startRow + n - 1;
        const address = `${startRow + 1}:${endRow + 1}`; // Excel 行号从 1 开始
        const insertRange = sheet.getRange(address);
        insertRange.insert(Excel.InsertShiftDirection.down);
      } else {
        // 插入当前列的单元格：从当前单元格往下推
        const col = active.columnIndex;
        const startRow = active.rowIndex;
        const insertRange = sheet.getRangeByIndexes(startRow, col, n, 1);
        insertRange.insert(Excel.InsertShiftDirection.down);
      }

      await context.sync();
    });

    showStatus(
      `✅ Inserted ${n} empty ${entireRow ? "row(s)" : "cell(s) in current column"} above selection.`
    );
  } catch (error) {
    showStatus("❌ " + (error.message || error));
    console.error(error);
  }
}

function showStatus(msg) {
  document.getElementById("status").innerText = msg;
}
