require("dotenv").config();
const axios = require("axios");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const numberToWordsRu = require("number-to-words-ru").convert;
const path = require("path");

const BEARER_TOKEN = process.env.BEARER_TOKEN;
const LAST_NAME = process.env.LAST_NAME || "";

const currentDate = new Date();
const startOfMonthDate = new Date(
  currentDate.getFullYear(),
  currentDate.getMonth() - 1,
  1
);
const lastDayOfMonthDate = new Date(
  currentDate.getFullYear(),
  currentDate.getMonth(),
  0
);
const currentMonthDate = new Date(
  currentDate.getFullYear(),
  currentDate.getMonth(),
  1
);

// Проверка существования токена авторизации
if (!BEARER_TOKEN) {
  console.error("BEARER_TOKEN is not defined in the environment variables");
  process.exit(1);
}

// Проверка наличия файла шаблона
const templatePath = "template.docx";
if (!fs.existsSync(templatePath)) {
  console.error(`Template file ${templatePath} not found`);
  process.exit(1);
}

async function getServerStatistics(date) {
  const response = await axios.get(
    `https://partners.cloud.vkplay.ru/api/v1/servers/statistic?date=${date}`,
    {
      headers: {
        Authorization: `Bearer ${BEARER_TOKEN}`,
      },
    }
  );
  return response.data;
}

async function calculateTotalSecondsByDay() {
  const dateString = `${startOfMonthDate.getFullYear()}-${String(
    startOfMonthDate.getMonth() + 1
  ).padStart(2, "0")}-${String(startOfMonthDate.getDate()).padStart(2, "0")}`;
  const statistics = await getServerStatistics(dateString);

  const serverDetails = statistics.map((server) => {
    const minutes = Math.floor(server.session_seconds / 60);
    const costPerMinute = parseFloat(server.playtime_cost);
    const earnings = minutes * costPerMinute;
    return {
      vm_name: server.vm_name,
      minutes,
      costPerMinute, // если понадобится в отчёте
      earnings,
    };
  });

  return { serverDetails };
}

function numberToWordsRuFormat(num) {
  try {
    return numberToWordsRu(num, {
      currency: "rub",
      declension: "nominative",
      showNumberParts: { integer: true, fractional: false },
      showCurrency: { integer: false, fractional: false },
    });
  } catch (error) {
    console.error("Error converting number to words:", error);
    throw error;
  }
}

function extractKopecks(amount) {
  return Math.round((amount - Math.floor(amount)) * 100);
}

function formatDate(date) {
  try {
    return date
      .toLocaleDateString("ru-RU", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      })
      .replace(/ г\./, "");
  } catch (error) {
    console.error("Error formatting date:", error);
    throw error;
  }
}

async function generateDocument(serverDetails, totalEarnings) {
  try {
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    let totalEarningsRub = Math.floor(totalEarnings);
    let totalEarningsKop = extractKopecks(totalEarnings);
    if (totalEarningsKop === 100) {
      totalEarningsRub += 1;
      totalEarningsKop = 0;
    }
    const totalEarningsText = `${totalEarningsRub} (${numberToWordsRuFormat(
      totalEarningsRub
    ).toLowerCase()}) руб. ${totalEarningsKop} коп.`;

    const data = {
      date: `${lastDayOfMonthDate.toLocaleString("ru-RU", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      })}`,
      startDate: formatDate(startOfMonthDate),
      endDate: formatDate(lastDayOfMonthDate),
      serverDetails: serverDetails.map((server, index) => ({
        index: index + 1,
        vm_name: server.vm_name,
        minutes: server.minutes,
        earnings: server.earnings.toFixed(2),
      })),
      totalEarnings: totalEarningsText,
    };

    console.log("Data to be rendered in the document:", data);

    doc.render(data);

    // Оптимизация содержимого перед сжатием
    zip.remove("word/settings.xml");

    const monthName = lastDayOfMonthDate.toLocaleString("ru-RU", {
      month: "long",
    });
    const year = lastDayOfMonthDate.getFullYear();
    const month = (lastDayOfMonthDate.getMonth() + 1)
      .toString()
      .padStart(2, "0");
    const fileName = `${year}-${month} (${monthName}) Акт выполненных работ ${LAST_NAME}.docx`;

    const outputDir = path.join(__dirname, "output");
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    const outputPath = path.join(outputDir, fileName);
    const buf = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
    fs.writeFileSync(outputPath, buf);

    console.log(`Document saved to ${outputPath}`);
  } catch (error) {
    console.error("Error generating document:", error.message);
    throw error;
  }
}

calculateTotalSecondsByDay()
  .then(({ serverDetails }) => {
    // Суммируем доход по каждому серверу
    const totalEarnings = serverDetails.reduce(
      (sum, server) => sum + server.earnings,
      0
    );

    generateDocument(serverDetails, totalEarnings)
      .then(() => {
        console.log("Документ успешно создан!");
      })
      .catch((error) => {
        console.error("Error generating document:", error.message);
      });

    const totalMinutes = serverDetails.reduce(
      (sum, server) => sum + server.minutes,
      0
    );

    console.log(
      `\nTotal gaming time in ${lastDayOfMonthDate.toLocaleString("ru-RU", {
        month: "long",
        year: "numeric",
      })}: ${totalMinutes} minutes.`
    );
    console.log(
      `Total money in ${lastDayOfMonthDate.toLocaleString("ru-RU", {
        month: "long",
        year: "numeric",
      })}: ${totalEarnings.toFixed(2)} rubles.`
    );
  })
  .catch((error) => {
    console.error("Error occurred:", error.message);
  });
