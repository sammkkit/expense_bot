const {Client,GatewayIntentBits} = require('discord.js');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

require('dotenv').config();

const client = new Client({
    intents:[
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent
    ]
});

client.once('ready',()=>{
    console.log(`logged in user ${client.user.name}`)
});

async function logToExcel(username, type, amount) {
    const file = path.join(__dirname, `${username}_expenses.xlsx`);
    const workbook = new ExcelJS.Workbook();
    let sheet;

    console.log(`===> Starting logToExcel for user: ${username}`);
    console.log(`Checking for file: ${file}`);

    if (fs.existsSync(file)) {
        console.log("✅ File exists. Reading workbook...");
        await workbook.xlsx.readFile(file);

        console.log("📄 Sheets found:");
        workbook.worksheets.forEach(ws => console.log(`   - ${ws.name}`));

        sheet = workbook.getWorksheet('Transactions');
        if (!sheet) {
            console.log("⚠️ 'Transactions' sheet missing. Creating new sheet...");
            sheet = workbook.addWorksheet('Transactions');
            sheet.columns = [
                { header: 'Date', key: 'date' },
                { header: 'Type', key: 'type' },
                { header: 'Amount', key: 'amount' }
            ];
        } else {
            console.log("✅ 'Transactions' worksheet loaded successfully.");
        }
    } else {
        console.log("🚫 File doesn't exist. Creating new workbook and sheet...");
        sheet = workbook.addWorksheet('Transactions');
        sheet.columns = [
            { header: 'Date', key: 'date' },
            { header: 'Type', key: 'type' },
            { header: 'Amount', key: 'amount' }
        ];
    }

    console.log(`📝 Adding new row: Date=${new Date().toLocaleString()}, Type=${type}, Amount=${amount}`);
    const date = new Date().toLocaleDateString();
    const row = sheet.addRow([date, type,amount]);

    console.log("📊 Current rows in worksheet (including header):");
    sheet.eachRow((row, rowNumber) => {
        console.log(`Row ${rowNumber}:`, row.values);
    });

    await workbook.xlsx.writeFile(file);
    console.log(`✅ Expense record saved to ${file}`);
}

client.on('messageCreate',async (message)=>{
    if(message.author.bot) return;
    const [command,amountStr] = message.content.trim().split(" ");
    const amount  = parseFloat(amountStr);
    const username = message.author.username;
    if(command == "!spend" && !isNaN(amount)){
        await logToExcel(username,"Expense",amount);
        message.reply(`💸 Logged **expense** of ${amount}`)
    } else if (command === "!gain" && !isNaN(amount)) {
        await logToExcel(username, "income", amount);
        message.reply(`💰 Logged **income** of ${amount}`);
    }else if (command === "!download") {
        const file = `${username}_expenses.xlsx`;
        console.log(`check 1 - ${file}`)
        if (fs.existsSync(file)) {
            await message.reply({ files: [file] });
        } else {
            message.reply("❌ No expense record found yet.");
        }
    }else if(command === "!commands"){
        message.reply("!spend, !gain, !download");
    }
})

client.login(process.env.TOKEN);
