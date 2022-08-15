const xlsx = require("xlsx");

const workbook = xlsx.readFile("./file.xls");

const columnA1ToR1C1 = (() => {
  const first = "A".charCodeAt(0);
  const last = "Z".charCodeAt(0);
  const length = last - first + 1;

  return (number, accumulator = "") => {
    if (number == 0) {
      return accumulator;
    }

    return columnA1ToR1C1(
      Math.floor((number - 1) / length),
      String.fromCharCode(((number - 1) % length) + first) + accumulator
    );
  };
})();

const cell = (row, column) => {
  return columnA1ToR1C1(column) + row;
};

const getData = (sheet, offset = 1) => {
  const posto = sheet(cell(1, offset + 8));
  const name = sheet(cell(2, offset + 1));

  let rows = [];
  let row_number = 6;
  while (true) {
    const data = sheet(cell(row_number, 1));

    if (!data) {
      break;
    }

    const nivel_res_lido = sheet(cell(row_number, offset + 1));
    const nivel_res_consol = sheet(cell(row_number, offset + 2));
    const vazao_defluente_lido = sheet(cell(row_number, offset + 3));
    const vazao_defluente_consol = sheet(cell(row_number, offset + 4));
    const vazao_afluente_lido = sheet(cell(row_number, offset + 5));
    const vazao_afluente_consol = sheet(cell(row_number, offset + 6));
    const vazao_increm_consol = sheet(cell(row_number, offset + 7));
    const vazao_natural_consol = sheet(cell(row_number, offset + 8));

    const row = {
      data,
      nivel_res_lido,
      nivel_res_consol,
      vazao_defluente_lido,
      vazao_defluente_consol,
      vazao_afluente_lido,
      vazao_afluente_consol,
      vazao_increm_consol,
      vazao_natural_consol,
    };

    rows.push(row);

    row_number += 1;
  }

  return {
    name,
    posto,
    rows,
  };
};

const items = [];

for (const name of workbook.SheetNames) {
  let offset = 1;

  const sheet = (cell) => workbook.Sheets[name][cell]?.w;

  while (true) {
    const data = getData(sheet, offset);

    offset += 8;

    if (!data.name) {
      break;
    }

    items.push({ sheet: name, ...data });
  }
}

require("fs").writeFileSync("result.json", JSON.stringify(items, null, 2));
