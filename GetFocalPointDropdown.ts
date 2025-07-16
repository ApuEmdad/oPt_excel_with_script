function main(workbook: ExcelScript.Workbook) {
  const sheetSectors = workbook.getWorksheet("Sectors");
  const sheetFocal = workbook.getWorksheet("focal_point");
  let sheetOutput = workbook.getWorksheet("sector-wise-focal");

  // Create or clear output sheet
  if (!sheetOutput) {
    sheetOutput = workbook.addWorksheet("sector-wise-focal");
  } else {
    sheetOutput.getUsedRange()?.clear();
  }

  // Step 1: Get sector list
  const sectorsData = sheetSectors.getUsedRange().getValues();
  const sectorHeader = sectorsData[0] as string[];
  const idxSector = sectorHeader.indexOf("Sector");
  if (idxSector === -1)
    throw new Error("Sector column not found in Sectors sheet.");

  const sectorList: string[] = [];
  for (let i = 1; i < sectorsData.length; i++) {
    const sector = sectorsData[i][idxSector] as string;
    if (sector && sector.trim()) {
      sectorList.push(sector);
    }
  }

  // Step 2: Get focal point mapping
  const focalData = sheetFocal.getUsedRange().getValues();
  const header = focalData[0] as string[];
  const colSector = header.indexOf("Sector");
  const colFocalPoint = header.indexOf("Focal Point");

  if (colSector === -1 || colFocalPoint === -1) {
    throw new Error("Missing 'Sector' or 'Focal Point' in focal_point sheet.");
  }

  const sectorMap = new Map<string, string[]>();

  for (let i = 1; i < focalData.length; i++) {
    const sector = focalData[i][colSector] as string;
    const focal = focalData[i][colFocalPoint] as string;
    if (!sector || !focal) continue;

    if (!sectorMap.has(sector)) {
      sectorMap.set(sector, []);
    }
    const list = sectorMap.get(sector)!;
    if (!list.includes(focal)) {
      list.push(focal);
    }
  }

  // Step 3: Write columns and define named ranges
  for (let col = 0; col < sectorList.length; col++) {
    const sector = sectorList[col];
    const focalList = sectorMap.get(sector) ?? [];

    // Column header
    sheetOutput.getCell(0, col).setValue(sector);

    // Write focal points
    for (let row = 0; row < focalList.length; row++) {
      sheetOutput.getCell(row + 1, col).setValue(focalList[row]);
    }

    // Clean name: replace space and slash
    const cleanName = sector.replace(/[ /]/g, "_");

    // Define named range (below header)
    if (focalList.length > 0) {
      const namedRange = sheetOutput.getRangeByIndexes(
        1,
        col,
        focalList.length,
        1
      );

      // Remove existing name if any
      const allNames = workbook.getNames();
      for (let n = 0; n < allNames.length; n++) {
        if (allNames[n].getName() === cleanName) {
          allNames[n].delete();
          break;
        }
      }

      workbook.addNamedItem(cleanName, namedRange);
    }
  }
}
