// import docx-templates
const { TemplateHandler, MimeType } = require("easy-template-x");
const fs = require("fs");
const xlsx = require("xlsx");

(async () => {
  const workbook = xlsx.readFile("data.xlsx");
  const sheet_name_list = workbook.SheetNames;
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  data.forEach(async (item) => {
    const template = fs.readFileSync("dump.docx");
    const handler = new TemplateHandler();
    const image = await fs.readFileSync(`./gambar/${item?.no}.png`);
    // scale image height and width to 50% of original size

    const data = {
      ...item,
      nama_lengkap_instansi: item?.nama_lengkap_instansi?.toUpperCase(),
      logo: {
        _type: "image",
        source: image,
        format: MimeType.Png,
        width: 60,
        height: 80,
        // scale to 50% of original size
      },
    };

    const doc = await handler.process(template, data);

    await fs.writeFileSync(`./output/${item.instansi}.docx`, doc);
  });
})();
