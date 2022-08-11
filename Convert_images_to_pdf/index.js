PDFDocument = require('pdfkit');
fs = require('fs');
//doc = new PDFDocument();
// Путь к папкам (Заносим все папки с jpg)
const path = "C:\\Users\\Vergel\\Desktop\\doc\\";
// Получаем наименование всех файлов
const directory = fs.readdirSync(path, { withFileTypes: true })
    .filter(d => d.isDirectory())
    .map(d => d.name);

// проходимся по папкам ФИО
directory.forEach(element => {
    fs.readdir(path + element, (err, files) => {
        doc = new PDFDocument();
        // Создаём файл pdf с наименованием ФИО папки
        doc.pipe(fs.createWriteStream(`${element}.pdf`))
        // Проходимся по всем файлам внутри папки
        for (let index = 0; index < files.length; ++index) {
            const f = files[index];
            // Если первый файл , то записываем на первую страницу
            if(index === 0) {
                // console.log(`${path + element}\\${f}`);
                doc.image(`${path + element}\\${f}`, {
                    fit: [500, 400],
                    align: 'center',
                    valign: 'center'
                 });
            }
            // В противном случаем необходимо создать addPage() и внести jpg
           else {
                doc.addPage().image(`${path + element}\\${f}`, {
                    fit: [500, 400],
                    align: 'center',
                    valign: 'center'
                 });
                 
            }
        }
        doc.end()
    });
});