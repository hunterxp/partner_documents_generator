require("dotenv").config();
const axios = require('axios');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const numberToWordsRu = require('number-to-words-ru').convert;
const path = require('path');

const BEARER_TOKEN = process.env.BEARER_TOKEN;
const LAST_NAME = process.env.LAST_NAME || '';

const currentDate = new Date();
const startOfMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
const lastDayOfMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 0);

const currentMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);

// Проверка существования токена авторизации
if (!BEARER_TOKEN) {
    console.error('BEARER_TOKEN is not defined in the environment variables');
    process.exit(1);
}

// Проверка наличия файла шаблона
const templatePath = 'template.docx';
if (!fs.existsSync(templatePath)) {
    console.error(`Template file ${templatePath} not found`);
    process.exit(1);
}

async function getServerStatistics(date) {
    const response = await axios.get(`https://partners.cloud.vkplay.ru/api/v1/servers/statistic?date=${date}`, {
        headers: {
            'Authorization': `Bearer ${BEARER_TOKEN}`
        }
    });
    return response.data;
}

async function calculateTotalSecondsByDay() {
    const dateString = currentMonthDate.toISOString().split('T')[0]; // Форматируем дату как YYYY-MM-DD
    const statistics = await getServerStatistics(dateString);

    let serverDetails = statistics.map(server => ({
        vm_name: server.vm_name,
        minutes: Math.floor(server.session_seconds / 60),
        earnings: (server.session_seconds / 60) * 0.3
    }));

    return { serverDetails };
}

function numberToWordsRuFormat(num) {
    try {
        const words = numberToWordsRu(num, {
            currency: 'rub',
            declension: 'nominative',
            showNumberParts: {
                integer: true,
                fractional: false,
            },
            showCurrency: {
                integer: false,
                fractional: false,
            },
        });
        return words;
    } catch (error) {
        console.error('Error converting number to words:', error);
        throw error;
    }
}

function extractKopecks(amount) {
    return Math.round((amount - Math.floor(amount)) * 100);
}

function formatDate(date) {
    try {
        const formattedDate = date.toLocaleDateString('ru-RU', {
            day: '2-digit',
            month: 'long',
            year: 'numeric'
        }).replace(/ г\./, '');  // Удаление " г." из строки
        return formattedDate;
    } catch (error) {
        console.error('Error formatting date:', error);
        throw error;
    }
}

async function generateDocument(serverDetails, totalEarnings) {
    try {
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        const totalEarningsRub = Math.floor(totalEarnings);
        const totalEarningsKop = extractKopecks(totalEarnings);
        const totalEarningsText = `${totalEarningsRub} (${(numberToWordsRuFormat(totalEarningsRub)).toLowerCase()}) руб. ${totalEarningsKop} коп.`;

        const data = {
            date: `${lastDayOfMonthDate.toLocaleString('ru-RU', { day: '2-digit', month: 'long', year: 'numeric' })}`,
            startDate: formatDate(startOfMonthDate),
            endDate: formatDate(lastDayOfMonthDate),
            serverDetails: serverDetails.map((server, index) => ({
                index: index + 1,
                vm_name: server.vm_name,
                minutes: server.minutes,
                earnings: server.earnings.toFixed(2)
            })),
            totalEarnings: totalEarningsText
        };

        console.log('Data to be rendered in the document:', data);

        doc.render(data);

        // Оптимизация содержимого перед сжатием
        zip.remove('word/settings.xml');

        const monthName = lastDayOfMonthDate.toLocaleString('ru-RU', { month: 'long' });
        const year = lastDayOfMonthDate.getFullYear();
        const month = lastDayOfMonthDate.getMonth() + 1; // January is 0
        const fileName = `${year}-${month.toString().padStart(2, '0')} (${monthName}) Акт выполненных работ ${LAST_NAME}.docx`;

        // Создаем директорию /output, если она не существует
        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        // Записываем файл в директорию /output
        const outputPath = path.join(outputDir, fileName);
        const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }); // добавим сжатие для оптимизации размера файла
        fs.writeFileSync(outputPath, buf);

        console.log(`Document saved to ${outputPath}`);
    } catch (error) {
        console.error('Error generating document:', error.message);
        throw error;
    }
}

calculateTotalSecondsByDay().then(({ serverDetails }) => {
    let grandTotal = 0;
    serverDetails.forEach(server => {
        grandTotal += server.minutes * 60; // конвертируем обратно в секунды для вычисления общего времени
    });

    const totalEarnings = grandTotal / 60 * 0.3;

    generateDocument(serverDetails, totalEarnings).then(() => {
        console.log("Документ успешно создан!");
    }).catch(error => {
        console.error('Error generating document:', error.message);
    });

    console.log(`\nTotal gaming time in ${lastDayOfMonthDate.toLocaleString('ru-RU', { month: 'long', year: 'numeric' })}: ${(grandTotal / 60).toFixed(0)} minutes.`);
    console.log(`Total money in ${lastDayOfMonthDate.toLocaleString('ru-RU', { month: 'long', year: 'numeric' })}: ${totalEarnings.toFixed(2)} rubles.`);
}).catch(error => {
    console.error('Error occurred:', error.message);
});
