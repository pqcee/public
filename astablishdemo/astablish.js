/**
 * @license
 * Author: pQCee
 * Description : AStablish implementation in Office Add-ins for Excel
 *
 * Copyright pQCee 2023. All rights reserved
 *
 * “Commons Clause” License Condition v1.0
 *
 * The Software is provided to you by the Licensor under the License, as defined
 * below, subject to the following condition.
 *
 * Without limiting other conditions in the License, the grant of rights under
 * the License will not include, and the License does not grant to you, the
 * right to Sell the Software.
 *
 * For purposes of the foregoing, “Sell” means practicing any or all of the
 * rights granted to you under the License to provide to third parties, for a
 * fee or other consideration (including without limitation fees for hosting or
 * consulting/ support services related to the Software), a product or service
 * whose value derives, entirely or substantially, from the functionality of the
 * Software. Any license notice or attribution required by the License must also
 * include this Commons Clause License Condition notice.
 *
 * Software: AStablish Office Add-in
 *
 * License: MIT License
 *
 * Licensor: pQCee Pte Ltd
 */

// =======================================
// REGISTER EVENTS FOR HTML GUI COMPONENTS
// =======================================

Office.onReady((info) => {
  // Check that we loaded into Excel
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btnCreateTable").onclick = createTable;
    document.getElementById("btnValidateWalletAddress").onclick = validateAddr;
    document.getElementById("btnCreateMessage").onclick = createMessage;
  }
});

// =================================
// GLOBAL - ASTABLISH BUNDLE EXPORTS
// =================================
const Buffer = astablishBundle.Buffer;
const ECPairFactory = astablishBundle.ECPair.ECPairFactory;
const bitcoinjs = astablishBundle.bitcoinjs;
const secp256k1 = astablishBundle.secp256k1;

// ==========================================
// GLOBAL - AUDIT TEMPLATE WORKSHEET SETTINGS
// ==========================================

/** Grey colour for cell shading */
const GREY = "#A5A5A5";

/** Bright yellow colour for cell shading */
const YELLOW = "#FFFF00";

/** Left-most column of the audit worksheet template */
const WS_LEFT_COLUMN = "A";

/** Top row of the audit worksheet template */
const WS_TOP_ROW = 1;

/** Top-left cell of the audit worksheet template (A1) */
const WS_START_CELL = "".concat(WS_LEFT_COLUMN, WS_TOP_ROW);

/** Right-most column of the audit worksheet template */
const WS_RIGHT_COLUMN = "H";

/** Content for Instructions Table */
const INSTRUCTIONS = [
  ["Instructions:"],
  ["1. Auditor fills up Message Params and send workbook to client."],
  ["2. Client choose BTC/ETH in Crypto column."],
  ["3. Client fills up Wallet Address & Public Key."],
  ["4. Client sign Message and fills up Digital Signature."],
  ["5. Client sends workbook back to Auditor."],
  ["6. Auditor clicks Validate button to verify wallet ownership."],
];

/** Top-left cell of Instructions Table */
const I_TABLE_START_CELL = WS_START_CELL;

/** Bottom-right cell of Instructions Table */
const I_TABLE_END_CELL = "".concat(WS_LEFT_COLUMN, INSTRUCTIONS.length);

/** Cell range of Instructions Table */
const I_TABLE_RANGE = "".concat(I_TABLE_START_CELL, ":", I_TABLE_END_CELL);

/** Number of spacer rows from top of worksheet to start of Main Table.
 *  It is computed from adding one empty row after the Instructions Table.
 */
const M_TABLE_SPACER_ROWS = INSTRUCTIONS.length + 1;

/** Default minimum number of data rows in Main Table */
const M_TABLE_DEFAULT_DATA_ROWS = 10;

/** Top row (header row) of Main Table */
const M_TABLE_TOP_ROW = WS_TOP_ROW + M_TABLE_SPACER_ROWS;

/** 2nd row (first row of data) of Main Table */
const M_TABLE_2ND_ROW = M_TABLE_TOP_ROW + 1;

/** Left-most column of Main Table*/
const M_TABLE_LEFT_COLUMN = WS_LEFT_COLUMN;

/** Top-left cell of Main Table */
const M_TABLE_START_CELL = "".concat(M_TABLE_LEFT_COLUMN, M_TABLE_TOP_ROW);

/** Right-most column of Main Table */
const M_TABLE_RIGHT_COLUMN = WS_RIGHT_COLUMN;

/** Zero-indexed value for Crypto column of Main Table */
//const M_TABLE_CRYPTO_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 1 - 1;

/** Zero-indexed value for Wallet Address column of Main Table */
const M_TABLE_ADDR_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 2 - 1;

/** Zero-indexed value for Public Key column of Main Table */
const M_TABLE_PUBKEY_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 3 - 1;

/** Zero-indexed value for Message column of Main Table */
//const M_TABLE_MSG_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 4 - 1;

/** Zero-indexed value for Digital Signature column of Main Table */
//const M_TABLE_SIG_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 5 - 1;

/** Zero-indexed value for Valid Wallet Address column of Main Table */
const M_TABLE_VALID_WALLET_COL = convertColToInt(M_TABLE_LEFT_COLUMN) + 6 - 1;

// =============
// BUTTON EVENTS
// =============

function createTable() {
  Excel.run((context) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    setupAuditTableTemplate(selectedSheet, M_TABLE_DEFAULT_DATA_ROWS);
    return context.sync();
  });
}

function validateAddr() {
  Excel.run((context) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

    // Derive cell range of data section in Main Table
    const DATA_START_CELL = "".concat(M_TABLE_LEFT_COLUMN, M_TABLE_2ND_ROW);
    let intDataRows = M_TABLE_DEFAULT_DATA_ROWS; // Placeholder for user input
    const DATA_ROWS = Math.max(intDataRows, M_TABLE_DEFAULT_DATA_ROWS);
    const BOTTOM_ROW = M_TABLE_2ND_ROW + DATA_ROWS - 1;
    const DATA_END_CELL = "".concat(M_TABLE_RIGHT_COLUMN, BOTTOM_ROW);
    const DATA_RANGE = "".concat(DATA_START_CELL, ":", DATA_END_CELL);

    // Load data portion of Main Table to proxy object in Office JS
    let objDataRange = selectedSheet.getRange(DATA_RANGE);
    objDataRange.load("values");
    return context.sync().then(() => {
      validateWalletAddress(objDataRange);
      return context.sync();
    });
  });
}

function createMessage() {
  Excel.run((context) => {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO: (1) Check/initialise MSG PARAMS Table, (2) Generate Message in Main Table

    return context.sync();
  });
}

// ================
// HELPER FUNCTIONS
// ================

/**
 * Convert Excel column alphabet to column number.
 *
 * @param {string} charColAlpha - Single character containing column alphabet.
 * @returns {number} Number equivalent of column alphabet, where A = 1, B = 2, etc.
 */
function convertColToInt(charColAlpha) {
  // This function does not receive arbitrary input from user.
  // Safe to assume the developer for this code will not pass in:
  // - zero-length string
  // - non-alphabet character
  // - double-alphabet string
  const ASCII_UPPER_CASE_A = "A".charCodeAt(0);
  return charColAlpha.toUpperCase().charCodeAt(0) - ASCII_UPPER_CASE_A + 1;
}

/**
 * Convert column number to Excel column alphabet.
 *
 * @param {number} intColNumber - Integer value of column number.
 * @returns {string} Character equivalent of column number, where 1 = A, 2 = B, etc.
 */
function convertIntToCol(intColNumber) {
  // This function does not receive arbitrary input from user.
  // Similar to convertColToInt(), safe to assume developer does not pass in invalid values.
  const ASCII_UPPER_CASE_A = "A".charCodeAt(0);
  return String.fromCharCode(intColNumber + ASCII_UPPER_CASE_A - 1);
}

/**
 * Set up the worksheet table layout for entering audit data.
 *
 * @param {Excel.Worksheet} objWS - Target worksheet to process.
 * @param {number} intDataRows - Number of rows of audit data to be filled in Main Table.
 */
function setupAuditTableTemplate(objWS, intDataRows) {
  /**
   * Converts value for Excel row height or column width from pixel to font
   * point. The conversion is achieved by using an approximation, where font
   * point = pixel * 0.75.
   *
   * @inner
   * @param {number} intPixelSize - Value for Excel row height or column width in pixels.
   * @returns {number} Equivalent width (floating point) value in Excel font points.
   */
  function pixelToPoint(intPixelSize) {
    return intPixelSize * 0.75;
  }

  /**
   * Apply continuous border lines to all sides of target cells.
   *
   * @inner
   * @param {Excel.Range} objRange - Range object containing target cells.
   */
  function addBorderLines(objRange) {
    const objRangeBorderCollection = objRange.format.borders;
    const STR_LINE = "Continuous";

    // Apply continuous lines to all sides of target cells
    objRangeBorderCollection.getItem("EdgeTop").style = STR_LINE;
    objRangeBorderCollection.getItem("EdgeRight").style = STR_LINE;
    objRangeBorderCollection.getItem("EdgeLeft").style = STR_LINE;
    objRangeBorderCollection.getItem("EdgeBottom").style = STR_LINE;
    objRangeBorderCollection.getItem("InsideHorizontal").style = STR_LINE;
    objRangeBorderCollection.getItem("InsideVertical").style = STR_LINE;
  }

  //
  // Calculate cell range of Main Table
  //

  /** Number of data rows in Main Table, has a minimum of 10 or more rows */
  const M_TABLE_DATA_ROWS = Math.max(intDataRows, M_TABLE_DEFAULT_DATA_ROWS);

  /** Bottom row of Main Table */
  const M_TABLE_BOTTOM_ROW = M_TABLE_TOP_ROW + M_TABLE_DATA_ROWS;

  /** Bottom-right cell of Main table */
  const M_TABLE_END_CELL = "".concat(M_TABLE_RIGHT_COLUMN, M_TABLE_BOTTOM_ROW);

  /** String containing cell range of Main Table */
  const M_TABLE_RANGE = "".concat(M_TABLE_START_CELL, ":", M_TABLE_END_CELL);

  //
  // Populate Instructions Table
  //
  objWS.getRange(I_TABLE_RANGE).values = INSTRUCTIONS;

  // =============================
  // MAIN TABLE BELOW INSTRUCTIONS
  // =============================
  // MAIN TABLE: HEADER
  const M_TABLE_HEADER = [
    [
      "No.",
      "Crypto",
      "Wallet Address",
      "Public Key",
      "Message",
      "Digital Signature",
      "Valid Wallet",
      "Verified",
    ],
  ];
  const M_TABLE_START_HDR = M_TABLE_START_CELL;
  const M_TABLE_END_HDR = "".concat(M_TABLE_RIGHT_COLUMN, M_TABLE_TOP_ROW);
  const M_TABLE_HDR_RANGE = "".concat(M_TABLE_START_HDR, ":", M_TABLE_END_HDR);
  objWS.getRange(M_TABLE_HDR_RANGE).values = M_TABLE_HEADER;
  objWS.getRange(M_TABLE_HDR_RANGE).format.font.bold = true;
  const M_TABLE_HDR_COLOURS = [
    GREY,
    YELLOW,
    YELLOW,
    YELLOW,
    GREY,
    YELLOW,
    GREY,
    GREY,
  ];

  for (
    let col = convertColToInt(M_TABLE_LEFT_COLUMN), row = M_TABLE_TOP_ROW;
    col <= convertColToInt(M_TABLE_RIGHT_COLUMN);
    col++
  ) {
    objWS.getRange("".concat(convertIntToCol(col), row)).format.fill.color =
      M_TABLE_HDR_COLOURS[col - 1];
  }

  // MAIN TABLE
  addBorderLines(objWS.getRange(M_TABLE_RANGE));
  objWS.getRange(M_TABLE_RANGE).format.horizontalAlignment = "Center";
  objWS.getRange(M_TABLE_RANGE).numberFormat = "0";

  // MAIN TABLE: DATA
  const M_TABLE_START_DAT = "".concat(M_TABLE_LEFT_COLUMN, M_TABLE_2ND_ROW);
  const M_TABLE_END_DAT = M_TABLE_END_CELL;
  const M_TABLE_DAT_RANGE = "".concat(M_TABLE_START_DAT, ":", M_TABLE_END_DAT);
  objWS.getRange(M_TABLE_DAT_RANGE).numberFormat = "@";

  // MAIN TABLE: DATA apply word-wrap from columns C to F
  const M_TABLE_START_WW = "".concat("C", M_TABLE_2ND_ROW);
  const M_TABLE_END_WW = "".concat("F", M_TABLE_BOTTOM_ROW);
  const M_TABLE_WW_RANGE = "".concat(M_TABLE_START_WW, ":", M_TABLE_END_WW);
  objWS.getRange(M_TABLE_WW_RANGE).format.wrapText = true;

  // MAIN TABLE: DATA left 2nd Column (Crypto)
  const M_TABLE_START_CC = "".concat("B", M_TABLE_2ND_ROW);
  const M_TABLE_END_CC = "".concat("B", M_TABLE_BOTTOM_ROW);
  const M_TABLE_CC_RANGE = "".concat(M_TABLE_START_CC, ":", M_TABLE_END_CC);
  objWS.getRange(M_TABLE_CC_RANGE).dataValidation.clear();
  objWS.getRange(M_TABLE_CC_RANGE).dataValidation.rule = {
    list: { inCellDropDown: true, source: "BTC,ETH" },
  };

  // MAIN TABLE: VALIDATION COLUMNS (Right 2 columns)
  const M_TABLE_START_VAL = "".concat("G", M_TABLE_2ND_ROW);
  const M_TABLE_END_VAL = M_TABLE_END_CELL;
  const M_TABLE_VAL_RANGE = "".concat(M_TABLE_START_VAL, ":", M_TABLE_END_VAL);
  objWS.getRange(M_TABLE_VAL_RANGE).conditionalFormats.clearAll();
  const trueConditionalFormat = objWS
    .getRange(M_TABLE_VAL_RANGE)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  trueConditionalFormat.textComparison.format.font.color = "#006100";
  trueConditionalFormat.textComparison.format.fill.color = "#C6EFCE";
  trueConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "TRUE",
  };
  const falseConditionalFormat = objWS
    .getRange(M_TABLE_VAL_RANGE)
    .conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  falseConditionalFormat.textComparison.format.font.color = "#9C0006";
  falseConditionalFormat.textComparison.format.fill.color = "#FFC7CE";
  falseConditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "FALSE",
  };

  // MAIN TABLE: Fill index column
  for (let i = 1; i <= intDataRows; i++) {
    let cell = "".concat(M_TABLE_LEFT_COLUMN, M_TABLE_TOP_ROW + i);
    objWS.getRange(cell).values = [[i.toString()]];
  }

  // ==============================================
  // MESSAGE PARAMS TABLE AT TOP-RIGHT OF WORKSHEET
  // ==============================================
  const P_TABLE_RIGHT_COLUMN = WS_RIGHT_COLUMN;
  const P_TABLE_TOP_ROW = WS_TOP_ROW + 1;
  const P_TABLE_LEFT_COLUMN = "G";
  const P_TABLE_START_HDR = "".concat(P_TABLE_LEFT_COLUMN, P_TABLE_TOP_ROW);
  const P_TABLE_END_HDR = "".concat(P_TABLE_RIGHT_COLUMN, P_TABLE_TOP_ROW);
  const P_TABLE_HDR_RANGE = "".concat(P_TABLE_START_HDR, ":", P_TABLE_END_HDR);
  objWS.getRange(P_TABLE_START_HDR).values = [["Message Params"]];
  objWS.getRange(P_TABLE_HDR_RANGE).merge(false);
  objWS.getRange(P_TABLE_HDR_RANGE).format.fill.color = YELLOW;
  objWS.getRange(P_TABLE_HDR_RANGE).format.font.bold = true;
  objWS.getRange(P_TABLE_HDR_RANGE).format.horizontalAlignment = "Center";
  const MSG_PARAMS = [["Seq. No."], ["Client Name"], ["Audit Date"]];
  objWS.getRange("G3:G5").values = MSG_PARAMS;
  addBorderLines(objWS.getRange("G2:H5"));
  objWS.getRange("H3:H5").numberFormat = "@";

  // =====================================
  // AUDIT COMMENTS TABLE BELOW MAIN TABLE
  // =====================================
  // AUDIT COMMENTS TABLE: HEADER
  const C_TABLE_TOP_ROW = M_TABLE_BOTTOM_ROW + 2;
  const C_TABLE_BOTTOM_ROW = C_TABLE_TOP_ROW + 10;
  const C_TABLE_LEFT_COLUMN = M_TABLE_LEFT_COLUMN;
  const C_TABLE_RIGHT_COLUMN = M_TABLE_RIGHT_COLUMN;
  const C_TABLE_START_HDR = "".concat(C_TABLE_LEFT_COLUMN, C_TABLE_TOP_ROW);
  const C_TABLE_END_HDR = "".concat(C_TABLE_RIGHT_COLUMN, C_TABLE_TOP_ROW);
  const C_TABLE_HDR_RANGE = "".concat(C_TABLE_START_HDR, ":", C_TABLE_END_HDR);
  objWS.getRange(C_TABLE_HDR_RANGE).merge(false);
  addBorderLines(objWS.getRange(C_TABLE_HDR_RANGE));
  objWS.getRange(C_TABLE_HDR_RANGE).format.horizontalAlignment = "Left";
  objWS.getRange(C_TABLE_HDR_RANGE).format.font.bold = true;
  objWS.getRange(C_TABLE_HDR_RANGE).format.fill.color = YELLOW;
  objWS.getRange(C_TABLE_START_HDR).values = [["Audit Comments"]];

  // AUDIT COMMENTS TABLE: DATA
  const C_TABLE_START_DAT = "".concat(C_TABLE_LEFT_COLUMN, C_TABLE_TOP_ROW + 1);
  const C_TABLE_END_DAT = "".concat(C_TABLE_RIGHT_COLUMN, C_TABLE_BOTTOM_ROW);
  const C_TABLE_DAT_RANGE = "".concat(C_TABLE_START_DAT, ":", C_TABLE_END_DAT);
  objWS.getRange(C_TABLE_DAT_RANGE).merge(false);
  addBorderLines(objWS.getRange(C_TABLE_DAT_RANGE));
  objWS.getRange(C_TABLE_DAT_RANGE).format.horizontalAlignment = "Left";

  // ================================
  // WORKSHEET RANGE FORMAT SETTINGS
  // ===============================
  /* TODO: optimise using set() method
  // ALSO GOOD: Use a "set" method to immediately set all the properties
  // without even needing to create a variable!
  worksheet.getRange("A1").set({
  numberFormat: [["0.00%"]],
  values: [[1]],
  format: {
      fill: {
          color: "red"
      }
  }
  });
  */
  const WS_BOTTOM_ROW = C_TABLE_BOTTOM_ROW;
  const WS_END_CELL = "".concat(WS_RIGHT_COLUMN, WS_BOTTOM_ROW);
  const WS_RANGE = "".concat(WS_START_CELL, ":", WS_END_CELL);
  const objWorkingRangeFormat = objWS.getRange(WS_RANGE).format;
  objWorkingRangeFormat.font.color = "#000000";
  objWorkingRangeFormat.font.name = "Calibri";
  objWorkingRangeFormat.font.size = 10;
  objWorkingRangeFormat.verticalAlignment = "Center";
  // Only Audit Comments Table: DATA need to be Top-justified
  objWS.getRange(C_TABLE_DAT_RANGE).format.verticalAlignment = "Top";

  // =======================================
  // WORKSHEET COLUMN WIDTHS AND ROW HEIGHTS
  // =======================================
  objWS.getRange("A1").format.columnWidth = pixelToPoint(29);
  objWS.getRange("B1").format.columnWidth = pixelToPoint(44);
  objWS.getRange("C1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("D1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("E1").format.columnWidth = pixelToPoint(138);
  objWS.getRange("F1").format.columnWidth = pixelToPoint(265);
  objWS.getRange("G1").format.columnWidth = pixelToPoint(74);
  objWS.getRange("H1").format.columnWidth = pixelToPoint(74);
  // Note: If you manually set the rowHeight, Excel no longer autofits rows
  //       to contents of cells with "wrapText = true". The way to do this
  //       is to not set the rowHeight programmatically.
  // objWS.getRange(WS_RANGE).format.rowHeight = pixelToPoint(17);
}

/**
 * Validate public key belongs to the wallet address in MAIN TABLE
 *
 * @param {Excel.Range} objDataRange - Cell range of data in Main Table.
 */
function validateWalletAddress(objDataRange) {
  // Create new array to populate updated data
  // I observed that Office JS context only updates the values back to the
  // Excel worksheet when a new array that contains entire range values in the
  // Excel.Range object are assigned to Excel.Range.values property.
  let data = objDataRange.values.map((arr) => arr.slice());

  for (let row = 0; row < data.length; row++) {
    let isWalletAddrFilled = data[row][M_TABLE_ADDR_COL] !== "";
    let isPublicKeyFilled = data[row][M_TABLE_PUBKEY_COL] !== "";

    // Check if supplied address === p2pkh(public key)
    if (isWalletAddrFilled && isPublicKeyFilled) {
      let pubkey = Buffer.from(data[row][M_TABLE_PUBKEY_COL], "hex");
      let { address } = bitcoinjs.payments.p2pkh({ pubkey });
      let isValidWallet = data[row][M_TABLE_ADDR_COL] === address;
      data[row][M_TABLE_VALID_WALLET_COL] = isValidWallet.toString();
    } else {
      // Do nothing and move on to next row
    }
  }

  // Update data in Main Table
  objDataRange.values = data;
}
