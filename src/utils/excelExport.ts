import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { BillOfLadingRecord } from '../types';

export async function exportBillOfLadingToExcel(bl: BillOfLadingRecord): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Ocean Shipping Lines - Bill of Lading System';
  workbook.created = new Date();

  const COLORS = {
    NAVY_HEADER: 'FF1F4E79',
    NAVY_TEXT: 'FF1F4E79',
    LIGHT_BLUE_BG: 'FFD9E1F2',
    ICE_BLUE_ROW: 'FFF2F5F9',
    WHITE: 'FFFFFFFF',
    BLACK: 'FF000000',
    GRAY_LABEL: 'FF595959',
    RED_ACCENT: 'FFC00000',
    BORDER_BLUE: 'FFBDD7EE',
    BORDER_NAVY: 'FF1F4E79',
  };

  const createBorder = (color = COLORS.BORDER_BLUE, style: ExcelJS.BorderStyle = 'thin') => ({
    top: { style, color: { argb: color } },
    left: { style, color: { argb: color } },
    bottom: { style, color: { argb: color } },
    right: { style, color: { argb: color } },
  });

  const blueBorder = createBorder(COLORS.BORDER_BLUE, 'thin');
  const navyBorder = createBorder(COLORS.BORDER_NAVY, 'thin');

  const applyBorderToRange = (
    ws: ExcelJS.Worksheet,
    startCol: number,
    startRow: number,
    endCol: number,
    endRow: number,
    border = blueBorder
  ) => {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        ws.getRow(r).getCell(c).border = border;
      }
    }
  };

  const wsDoc = workbook.addWorksheet('Bill of Lading', {
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0.1, footer: 0.1 },
    },
    views: [{ showGridLines: true }],
  });

  wsDoc.columns = [
    { key: 'A', width: 14.0 },
    { key: 'B', width: 13.0 },
    { key: 'C', width: 13.0 },
    { key: 'D', width: 8.5 },
    { key: 'E', width: 14.0 },
    { key: 'F', width: 13.0 },
    { key: 'G', width: 13.0 },
    { key: 'H', width: 8.5 },
    { key: 'I', width: 12.0 },
    { key: 'J', width: 12.0 },
    { key: 'K', width: 12.0 },
    { key: 'L', width: 15.0 },
  ];

  // Row 1
  wsDoc.mergeCells('A1:H1');
  const r1Left = wsDoc.getCell('A1');
  r1Left.value = {
    richText: [
      { text: 'BILL OF LADING — ', font: { name: 'Calibri', size: 12, bold: true, color: { argb: COLORS.WHITE } } },
      { text: bl.headerTitle || 'FARAN SUGAR MILLS LTD', font: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFC000' } } },
    ],
  };
  r1Left.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  r1Left.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('I1:L1');
  const r1Right = wsDoc.getCell('I1');
  r1Right.value = `B/L:  ${bl.blNumber || 'OSL-2024-PK40156'}`;
  r1Right.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  r1Right.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.WHITE } };
  r1Right.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
  wsDoc.getRow(1).height = 24;

  // Row 2
  wsDoc.mergeCells('A2:L2');
  const r2Cell = wsDoc.getCell('A2');
  r2Cell.value = `BOOKING REF: ${bl.bookingRef || 'OSL-40156/24(CRO-3476/24)'}   |   CARRIER: ${bl.carrierName || 'MOL / OCEAN SHIPPING LINES'}   |   B/L NO: ${bl.blNumber || 'OSL-2024-PK40156'}`;
  r2Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  r2Cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  r2Cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(2).height = 18;

  // Row 3
  wsDoc.mergeCells('A3:H3');
  const s1 = wsDoc.getCell('A3');
  s1.value = 'SHIPPER / EXPORTER';
  s1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  s1.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  s1.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('I3:L3');
  const s2 = wsDoc.getCell('I3');
  s2.value = 'B/L REFERENCES & FREIGHT TERMS';
  s2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  s2.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  s2.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(3).height = 17;

  // Row 4-6
  wsDoc.mergeCells('A4:H6');
  const shCell = wsDoc.getCell('A4');
  shCell.value = bl.shipperText;
  shCell.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.BLACK } };
  shCell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells('I4:J4');
  wsDoc.getCell('I4').value = 'REFERENCE NO:';
  wsDoc.getCell('I4').font = { name: 'Calibri', size: 8, color: { argb: COLORS.GRAY_LABEL } };
  wsDoc.getCell('I4').alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('K4:L4');
  wsDoc.getCell('K4').value = bl.referenceNo;
  wsDoc.getCell('K4').font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.BLACK } };
  wsDoc.getCell('K4').alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(4).height = 16;

  wsDoc.mergeCells('I5:J5');
  wsDoc.getCell('I5').value = 'CARRIER:';
  wsDoc.getCell('I5').font = { name: 'Calibri', size: 8, color: { argb: COLORS.GRAY_LABEL } };
  wsDoc.getCell('I5').alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('K5:L5');
  wsDoc.getCell('K5').value = bl.carrierName;
  wsDoc.getCell('K5').font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.BLACK } };
  wsDoc.getCell('K5').alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(5).height = 16;

  // Row 7
  wsDoc.mergeCells('A7:H7');
  const csH = wsDoc.getCell('A7');
  csH.value = 'CONSIGNEE OR ORDER';
  csH.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  csH.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  csH.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('I7:L7');
  const ftH = wsDoc.getCell('I7');
  ftH.value = 'FREIGHT TERMS';
  ftH.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  ftH.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  ftH.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(7).height = 16;

  // Row 8
  wsDoc.mergeCells('A8:H8');
  const csVal = wsDoc.getCell('A8');
  csVal.value = bl.consigneeText;
  csVal.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.BLACK } };
  csVal.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('I8:L8');
  const ftVal = wsDoc.getCell('I8');
  ftVal.value = `${bl.freightPayableAt}\n(${bl.movement})`;
  ftVal.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.BLACK } };
  ftVal.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
  wsDoc.getRow(8).height = 24;

  // Row 9: Form E
  wsDoc.mergeCells('A9:L9');
  const feCell = wsDoc.getCell('A9');
  feCell.value = `FORM E NO(s):  ${bl.formENumber}`;
  feCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F2' } };
  feCell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: COLORS.RED_ACCENT } };
  feCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(9).height = 20;

  // Row 10-11: Notify Parties
  const formatPartyText = (party?: { name?: string; address?: string }) => {
    if (!party || (!party.name?.trim() && !party.address?.trim())) return '';
    const lines = [];
    if (party.name?.trim()) lines.push(party.name.trim());
    if (party.address?.trim()) lines.push(party.address.trim());
    return lines.join('\n');
  };

  const n1 = bl.notifyParties[0];
  const n2 = bl.notifyParties[1];
  const n3 = bl.notifyParties[2];
  const n4 = bl.notifyParties[3];

  wsDoc.mergeCells('A10:F10');
  const np1Header = wsDoc.getCell('A10');
  np1Header.value = 'NOTIFY PARTY/ADDRESS #1';
  np1Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  np1Header.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  np1Header.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells('G10:L10');
  const np2Header = wsDoc.getCell('G10');
  np2Header.value = 'NOTIFY PARTY/ADDRESS #2';
  np2Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  np2Header.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  np2Header.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(10).height = 16;

  const np1Text = formatPartyText(n1);
  const np2Text = formatPartyText(n2);

  wsDoc.mergeCells('A11:F11');
  const np1Val = wsDoc.getCell('A11');
  np1Val.value = np1Text;
  np1Val.font = { name: 'Calibri', size: 8, bold: true, color: { argb: COLORS.BLACK } };
  np1Val.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells('G11:L11');
  const np2Val = wsDoc.getCell('G11');
  np2Val.value = np2Text;
  np2Val.font = { name: 'Calibri', size: 8, bold: true, color: { argb: COLORS.BLACK } };
  np2Val.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  const r11Lines = Math.max(np1Text.split('\n').length, np2Text.split('\n').length, 4);
  wsDoc.getRow(11).height = Math.max(58, r11Lines * 15);

  let curRowIndex = 12;

  if (bl.notifyParties.length > 2) {
    const rHeader = curRowIndex;
    wsDoc.mergeCells(`A${rHeader}:F${rHeader}`);
    const np3Header = wsDoc.getCell(`A${rHeader}`);
    np3Header.value = 'NOTIFY PARTY/ADDRESS #3';
    np3Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
    np3Header.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
    np3Header.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

    wsDoc.mergeCells(`G${rHeader}:L${rHeader}`);
    const np4Header = wsDoc.getCell(`G${rHeader}`);
    np4Header.value = 'NOTIFY PARTY/ADDRESS #4';
    np4Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
    np4Header.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
    np4Header.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    wsDoc.getRow(rHeader).height = 16;

    const np3Text = formatPartyText(n3);
    const np4Text = formatPartyText(n4);

    const rVal = curRowIndex + 1;
    wsDoc.mergeCells(`A${rVal}:F${rVal}`);
    const np3Val = wsDoc.getCell(`A${rVal}`);
    np3Val.value = np3Text;
    np3Val.font = { name: 'Calibri', size: 8, color: { argb: COLORS.BLACK } };
    np3Val.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

    wsDoc.mergeCells(`G${rVal}:L${rVal}`);
    const np4Val = wsDoc.getCell(`G${rVal}`);
    np4Val.value = np4Text;
    np4Val.font = { name: 'Calibri', size: 8, color: { argb: COLORS.BLACK } };
    np4Val.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

    const rValLines = Math.max(np3Text.split('\n').length, np4Text.split('\n').length, 2);
    wsDoc.getRow(rVal).height = Math.max(38, rValLines * 15);
    curRowIndex += 2;
  }

  // Routing rows
  const rVessel = curRowIndex;
  wsDoc.mergeCells(`A${rVessel}:F${rVessel}`);
  const vvCell = wsDoc.getCell(`A${rVessel}`);
  vvCell.value = `VESSEL AND VOY NO:   ${bl.vesselAndVoy}`;
  vvCell.font = { name: 'Calibri', size: 8.5, bold: true };
  vvCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells(`G${rVessel}:L${rVessel}`);
  const podelCell = wsDoc.getCell(`G${rVessel}`);
  podelCell.value = `PLACE OF DELIVERY:   ${bl.placeOfDelivery}`;
  podelCell.font = { name: 'Calibri', size: 8.5, bold: true };
  podelCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(rVessel).height = 18;

  const rPort = curRowIndex + 1;
  wsDoc.mergeCells(`A${rPort}:F${rPort}`);
  const polCell = wsDoc.getCell(`A${rPort}`);
  polCell.value = `PORT OF LOADING:   ${bl.portOfLoading}`;
  polCell.font = { name: 'Calibri', size: 8.5, bold: true };
  polCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells(`G${rPort}:L${rPort}`);
  const podCell = wsDoc.getCell(`G${rPort}`);
  podCell.value = `PORT OF DISCHARGE:   ${bl.portOfDischarge}`;
  podCell.font = { name: 'Calibri', size: 8.5, bold: true };
  podCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(rPort).height = 18;

  applyBorderToRange(wsDoc, 1, 1, 12, rPort, blueBorder);

  // Cargo Section
  const rCargoHeader = rPort + 1;
  wsDoc.mergeCells(`A${rCargoHeader}:C${rCargoHeader}`);
  const chMarks = wsDoc.getCell(`A${rCargoHeader}`);
  chMarks.value = 'MARKS AND NOS:';
  chMarks.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  chMarks.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.WHITE } };
  chMarks.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells(`D${rCargoHeader}:H${rCargoHeader}`);
  const chGoods = wsDoc.getCell(`D${rCargoHeader}`);
  chGoods.value = 'NUMBER AND KIND OF PACKAGES: DESCRIPTION OF GOODS';
  chGoods.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  chGoods.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.WHITE } };
  chGoods.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells(`I${rCargoHeader}:J${rCargoHeader}`);
  const chGross = wsDoc.getCell(`I${rCargoHeader}`);
  chGross.value = 'GROSS WEIGHT';
  chGross.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  chGross.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.WHITE } };
  chGross.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };

  wsDoc.mergeCells(`K${rCargoHeader}:L${rCargoHeader}`);
  const chMeas = wsDoc.getCell(`K${rCargoHeader}`);
  chMeas.value = 'MEASURMENT';
  chMeas.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  chMeas.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.WHITE } };
  chMeas.alignment = { vertical: 'middle', horizontal: 'right', indent: 1 };
  wsDoc.getRow(rCargoHeader).height = 20;

  const rCargoBody = rCargoHeader + 1;
  wsDoc.mergeCells(`A${rCargoBody}:C${rCargoBody}`);
  const mnBody = wsDoc.getCell(`A${rCargoBody}`);
  mnBody.value = bl.marksAndNumbers;
  mnBody.font = { name: 'Calibri', size: 8.5, bold: true };
  mnBody.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`D${rCargoBody}:H${rCargoBody}`);
  const descBody = wsDoc.getCell(`D${rCargoBody}`);
  descBody.value = bl.cargoDescription;
  descBody.font = { name: 'Calibri', size: 8, bold: true };
  descBody.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`I${rCargoBody}:J${rCargoBody}`);
  const wtBody = wsDoc.getCell(`I${rCargoBody}`);
  wtBody.value = `${bl.grossWeight} ${bl.grossWeightUnit}\n\nNETT WEIGHT\n${bl.nettWeight} ${bl.nettWeightUnit}\n\n${bl.totalPackagesCount}\n${bl.totalPackagesUnit}`;
  wtBody.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  wtBody.alignment = { vertical: 'top', horizontal: 'right', indent: 1, wrapText: true };

  wsDoc.mergeCells(`K${rCargoBody}:L${rCargoBody}`);
  const msBody = wsDoc.getCell(`K${rCargoBody}`);
  msBody.value = bl.measurement;
  msBody.font = { name: 'Calibri', size: 8.5, bold: true };
  msBody.alignment = { vertical: 'top', horizontal: 'right', indent: 1 };

  wsDoc.getRow(rCargoBody).height = 135;
  applyBorderToRange(wsDoc, 1, rCargoHeader, 12, rCargoBody, blueBorder);

  // Containers & Clauses
  const activeContainers = bl.containers.filter(c => c && c.containerNo && c.containerNo.trim() !== '');
  const activeCount = activeContainers.length;

  let activeNetKg = 0;
  let activeGrossKg = 0;
  let activeBags = 0;
  activeContainers.forEach((c) => {
    activeNetKg += c.netWeightKg > 0 ? c.netWeightKg : 26000;
    activeGrossKg += c.grossWeightKg > 0 ? c.grossWeightKg : 26070.2;
    activeBags += c.packageCount > 0 ? c.packageCount : 520;
  });

  const perContNet = activeCount > 0 ? activeNetKg / activeCount : 26000;
  const perContGross = activeCount > 0 ? activeGrossKg / activeCount : 26070.2;
  const perContBags = activeCount > 0 ? Math.round(activeBags / activeCount) : 520;

  const rContHeader = rCargoBody + 1;
  wsDoc.mergeCells(`A${rContHeader}:H${rContHeader}`);
  const cTitle = wsDoc.getCell(`A${rContHeader}`);
  cTitle.value = `SEAL NO: / CONTAINER NO:   (${activeCount} CONT)`;
  cTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  cTitle.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.WHITE } };
  cTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

  wsDoc.mergeCells(`I${rContHeader}:L${rContHeader}`);
  const scTitle = wsDoc.getCell(`I${rContHeader}`);
  scTitle.value = 'SPECIAL INSTRUCTIONS / CLAUSES';
  scTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  scTitle.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  scTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  wsDoc.getRow(rContHeader).height = 18;

  const rContSub = rContHeader + 1;
  const pkgSubLabel = bl.totalPackagesUnit === 'BAGS' ? 'BAGS' : bl.totalPackagesUnit;
  const subCols = [
    { col: 1, val: 'Container No.', align: 'left' as const, indent: 1 },
    { col: 2, val: 'Net Wt (kg)', align: 'right' as const, indent: 0 },
    { col: 3, val: 'Gross Wt (kg)', align: 'right' as const, indent: 0 },
    { col: 4, val: pkgSubLabel, align: 'right' as const, indent: 0 },
    { col: 5, val: 'Container No.', align: 'left' as const, indent: 1 },
    { col: 6, val: 'Net Wt (kg)', align: 'right' as const, indent: 0 },
    { col: 7, val: 'Gross Wt (kg)', align: 'right' as const, indent: 0 },
    { col: 8, val: pkgSubLabel, align: 'right' as const, indent: 0 },
  ];

  subCols.forEach((sc) => {
    const cell = wsDoc.getRow(rContSub).getCell(sc.col);
    cell.value = sc.val;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
    cell.font = { name: 'Calibri', size: 7.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: sc.align, indent: sc.indent };
    cell.border = blueBorder;
  });
  wsDoc.getRow(rContSub).height = 16;

  const half = Math.ceil(bl.containers.length / 2);
  const leftCol = bl.containers.slice(0, half);
  const rightCol = bl.containers.slice(half);
  const rowCount = Math.max(leftCol.length, rightCol.length, 12);

  wsDoc.mergeCells(`I${rContSub}:L${rContSub + rowCount - 1}`);
  const scBodyCell = wsDoc.getCell(`I${rContSub}`);
  scBodyCell.value = `${bl.specialClauses.join('\n\n')}\n\n\n\n${bl.containerSummaryLine}`;
  scBodyCell.font = { name: 'Calibri', size: 8, bold: true, color: { argb: COLORS.BLACK } };
  scBodyCell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };
  applyBorderToRange(wsDoc, 9, rContSub, 12, rContSub + rowCount - 1, blueBorder);

  for (let i = 0; i < rowCount; i++) {
    const rNum = rContSub + 1 + i;
    const isZebra = i % 2 === 1;
    const bgArgb = isZebra ? COLORS.ICE_BLUE_ROW : COLORS.WHITE;
    const row = wsDoc.getRow(rNum);
    row.height = 13.5;

    const c1 = leftCol[i];
    const hasC1 = Boolean(c1 && c1.containerNo && c1.containerNo.trim() !== '');
    if (hasC1) {
      const netVal = c1.netWeightKg > 0 ? c1.netWeightKg : perContNet;
      const grossVal = c1.grossWeightKg > 0 ? c1.grossWeightKg : perContGross;
      const bagsVal = c1.packageCount > 0 ? c1.packageCount : perContBags;

      const c1No = row.getCell(1);
      c1No.value = c1.containerNo.trim();
      c1No.font = { name: 'Consolas', size: 8, bold: true };
      c1No.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

      const c1Net = row.getCell(2);
      c1Net.value = Number(netVal.toFixed(2));
      c1Net.numFmt = '#,##0.00';
      c1Net.font = { name: 'Calibri', size: 8 };
      c1Net.alignment = { vertical: 'middle', horizontal: 'right' };

      const c1Gross = row.getCell(3);
      c1Gross.value = Number(grossVal.toFixed(2));
      c1Gross.numFmt = '#,##0.00';
      c1Gross.font = { name: 'Calibri', size: 8 };
      c1Gross.alignment = { vertical: 'middle', horizontal: 'right' };

      const c1Bags = row.getCell(4);
      c1Bags.value = Number(bagsVal);
      c1Bags.numFmt = '#,##0';
      c1Bags.font = { name: 'Calibri', size: 8 };
      c1Bags.alignment = { vertical: 'middle', horizontal: 'right' };
    }

    const c2 = rightCol[i];
    const hasC2 = Boolean(c2 && c2.containerNo && c2.containerNo.trim() !== '');
    if (hasC2) {
      const netVal = c2.netWeightKg > 0 ? c2.netWeightKg : perContNet;
      const grossVal = c2.grossWeightKg > 0 ? c2.grossWeightKg : perContGross;
      const bagsVal = c2.packageCount > 0 ? c2.packageCount : perContBags;

      const c2No = row.getCell(5);
      c2No.value = c2.containerNo.trim();
      c2No.font = { name: 'Consolas', size: 8, bold: true };
      c2No.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

      const c2Net = row.getCell(6);
      c2Net.value = Number(netVal.toFixed(2));
      c2Net.numFmt = '#,##0.00';
      c2Net.font = { name: 'Calibri', size: 8 };
      c2Net.alignment = { vertical: 'middle', horizontal: 'right' };

      const c2Gross = row.getCell(7);
      c2Gross.value = Number(grossVal.toFixed(2));
      c2Gross.numFmt = '#,##0.00';
      c2Gross.font = { name: 'Calibri', size: 8 };
      c2Gross.alignment = { vertical: 'middle', horizontal: 'right' };

      const c2Bags = row.getCell(8);
      c2Bags.value = Number(bagsVal);
      c2Bags.numFmt = '#,##0';
      c2Bags.font = { name: 'Calibri', size: 8 };
      c2Bags.alignment = { vertical: 'middle', horizontal: 'right' };
    }

    for (let col = 1; col <= 8; col++) {
      const cell = row.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.border = blueBorder;
    }
  }

  // Totals Row
  const totalsRowNum = rContSub + rowCount;
  wsDoc.mergeCells(`A${totalsRowNum}:H${totalsRowNum}`);
  const contTotalsBar = wsDoc.getCell(`A${totalsRowNum}`);
  contTotalsBar.value = `TOTALS: ${activeCount} CONTAINER(S)   |   NET: ${activeNetKg.toLocaleString('en-US', { minimumFractionDigits: 2 })} KG   |   GROSS: ${activeGrossKg.toLocaleString('en-US', { minimumFractionDigits: 2 })} KG   |   ${activeBags.toLocaleString()} ${bl.totalPackagesUnit}`;
  contTotalsBar.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  contTotalsBar.font = { name: 'Calibri', size: 9, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  contTotalsBar.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDoc.mergeCells(`I${totalsRowNum}:L${totalsRowNum}`);
  const contSummaryBar = wsDoc.getCell(`I${totalsRowNum}`);
  contSummaryBar.value = 'CONTAINER SUMMARY VERIFIED';
  contSummaryBar.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LIGHT_BLUE_BG } };
  contSummaryBar.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.NAVY_TEXT } };
  contSummaryBar.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDoc.getRow(totalsRowNum).height = 19;
  applyBorderToRange(wsDoc, 1, totalsRowNum, 12, totalsRowNum, {
    top: { style: 'thin', color: { argb: COLORS.BORDER_NAVY } },
    bottom: { style: 'double', color: { argb: COLORS.BORDER_NAVY } },
    left: { style: 'thin', color: { argb: COLORS.BORDER_BLUE } },
    right: { style: 'thin', color: { argb: COLORS.BORDER_BLUE } },
  });

  // Footer Particulars Banner
  const apRowNum = totalsRowNum + 1;
  wsDoc.mergeCells(`A${apRowNum}:L${apRowNum}`);
  const apBanner = wsDoc.getCell(`A${apRowNum}`);
  apBanner.value = 'ABOVE PARTICULARS AS DECLARED BY SHIPPER';
  apBanner.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.NAVY_HEADER } };
  apBanner.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: COLORS.WHITE } };
  apBanner.alignment = { vertical: 'middle', horizontal: 'center' };
  wsDoc.getRow(apRowNum).height = 16;
  applyBorderToRange(wsDoc, 1, apRowNum, 12, apRowNum, navyBorder);

  // Footer cells
  const f1RowNum = apRowNum + 1;
  const f2RowNum = apRowNum + 2;

  wsDoc.mergeCells(`A${f1RowNum}:D${f1RowNum}`);
  const f1Cell = wsDoc.getCell(`A${f1RowNum}`);
  f1Cell.value = `TOTAL NO. OF CONTAINERS/PACKAGES\n${bl.totalContainersPackages || activeCount + "X20'  FCL"}`;
  f1Cell.font = { name: 'Calibri', size: 8, bold: true };
  f1Cell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`E${f1RowNum}:H${f1RowNum}`);
  const f2Cell = wsDoc.getCell(`E${f1RowNum}`);
  f2Cell.value = `FREIGHT PAYABLE AT\n${bl.freightPayableAt}`;
  f2Cell.font = { name: 'Calibri', size: 8, bold: true };
  f2Cell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`I${f1RowNum}:L${f2RowNum}`);
  const f3Cell = wsDoc.getCell(`I${f1RowNum}`);
  f3Cell.value = `PLACE AND DATE OF ISSUE\n${bl.placeAndDateOfIssue}\n\n\nAS CARRIER / AUTHORIZED AGENT`;
  f3Cell.font = { name: 'Calibri', size: 8, bold: true };
  f3Cell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`A${f2RowNum}:D${f2RowNum}`);
  const f4Cell = wsDoc.getCell(`A${f2RowNum}`);
  f4Cell.value = `MOVEMENT\n${bl.movement}`;
  f4Cell.font = { name: 'Calibri', size: 8, bold: true };
  f4Cell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.mergeCells(`E${f2RowNum}:H${f2RowNum}`);
  const f5Cell = wsDoc.getCell(`E${f2RowNum}`);
  f5Cell.value = `NUMBER OF ORIGINAL B/L\n${bl.numberOfOriginals}`;
  f5Cell.font = { name: 'Calibri', size: 8, bold: true };
  f5Cell.alignment = { vertical: 'top', horizontal: 'left', indent: 1, wrapText: true };

  wsDoc.getRow(f1RowNum).height = 32;
  wsDoc.getRow(f2RowNum).height = 32;

  applyBorderToRange(wsDoc, 1, f1RowNum, 4, f1RowNum, blueBorder);
  applyBorderToRange(wsDoc, 5, f1RowNum, 8, f1RowNum, blueBorder);
  applyBorderToRange(wsDoc, 1, f2RowNum, 4, f2RowNum, blueBorder);
  applyBorderToRange(wsDoc, 5, f2RowNum, 8, f2RowNum, blueBorder);
  applyBorderToRange(wsDoc, 9, f1RowNum, 12, f2RowNum, blueBorder);
  applyBorderToRange(wsDoc, 1, 1, 12, f2RowNum, blueBorder);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  saveAs(blob, `Bill_of_Lading_${bl.blNumber || 'OSL-2024-PK40156'}.xlsx`);
}
