const puppeteer = require('puppeteer');
const {timeout} = require('../tools/tools.js');
const moment = require('moment');

const schedule = require('node-schedule');
const fsPromises = require("fs/promises");
const xlsx = require('node-xlsx');
const nodemailer = require('nodemailer');


const sheetPath = 'trd.xlsx';
let date = '';

// 抓取充电量
async function getWebpageData(browser) {
  try {
    const page = await browser.newPage();
    let status = await page.goto('https://www.teld.cn', {timeout: 60 * 3 * 1000});
    await timeout(3000);

    console.log('start!');
    let powerArr = await page.evaluate(() => {
      let as = [...document.querySelectorAll('.one-info:first-child .middle span:last-of-type:not(.space)')];

      return as.map((a) =>{
          return {
            text: a.innerText
          }
      });
    });

    const num = powerArr[0].text.replace(/,/gi, '')

    let powerNum = Number(num)
    console.log(date, powerNum)

    await page.close()
    await browser.close();

    return powerNum;
  } catch (error) {
    console.error(error)
    date = moment().format("YYYY-MM-DD HH:mm:ss")
    console.log('重试', date)
    return await getWebpageData(browser)
  }
}

// 写入xlsx
async function writeSheet(powerNum) {
  // 先查出历史数据
  const sheets = xlsx.parse(sheetPath);
      
  let beforeElsxData = []
  // 遍历 sheet
  sheets.forEach(function(sheet){
      // console.log(sheet['name']);
      // 读取每行内容
      for(const rowId in sheet['data']){
          const row=sheet['data'][rowId];
          // console.log(row);
          beforeElsxData.push(row)
      }
  });

  const lastRowData = beforeElsxData[beforeElsxData.length - 1]
  const subjectText = `时间：${date}, 充电量：${powerNum}, 增量：${powerNum - lastRowData[1]}`

  console.log(subjectText)

  beforeElsxData.push([
    date,
    powerNum + '',
    (powerNum - lastRowData[1]) + ''
  ])

  const data = [{
    name: 'sheet1',
    data: beforeElsxData
  }]
  const buffer = xlsx.build(data);

  // 写入文件
  const err = await fsPromises.writeFile(sheetPath, buffer)
  if (err) {
    console.log("Write failed: " + err);
    return;
  }

  console.log("Write completed.");
  return subjectText
}

// 发送邮件
async function sendEmail(subjectText) {
  // 开启一个 SMTP 连接池
  let transporter = nodemailer.createTransport({
    host: 'smtp.qq.com',
    secureConnection: true, // use SSL
    port: 465,
    secure: true, // secure:true for port 465, secure:false for port 587
    auth: {
        user: 'jiangyige1990@qq.com',
        pass: 'vcanaoaitnoocaji' // QQ邮箱需要使用授权码
    }
  });

  // 设置邮件内容（谁发送什么给谁）
  let mailOptions = {
      from: 'jiangyige1990@qq.com', // 发件人
      to: 'jiangyige1990@live.com', // 收件人
      subject: subjectText, // 主题
      text: 'TRD' + moment().format('YYYY-MM-DD') + '充电量', // plain text body
      html: '<b>详情看附件</b>', // html body
      // 下面是发送附件，不需要就注释掉
      attachments: [{
        filename: 'trd.xlsx',
        path: sheetPath
      }]
  };

  // 使用先前创建的传输器的 sendMail 方法传递消息对象
  transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        return console.log(error);
      }
      console.log(`Message: ${info.messageId}`);
      console.log(`sent: ${info.response}`);
  });
}


// 定时任务
const rule = new schedule.RecurrenceRule();
// runs at 10:00:00
rule.hour = 10;
rule.minute = 0;
rule.second = 0;
rule.tz = 'Asia/Shanghai';

const job = schedule.scheduleJob(rule, () => {
  console.log('job schedule', new Date().toLocaleString());

  puppeteer.launch({args: ['--no-sandbox', '--disable-setuid-sandbox']}).then(async browser => {
  date = moment().format("YYYY-MM-DD HH:mm:ss")
  // 抓取充电量
  const powerNum = await getWebpageData(browser);
  if (!powerNum) {
    console.log('抓取充电量失败')
    return;
  }

  // 写入xlsx
  const subjectText = await writeSheet(powerNum);
  if (!subjectText) {
    console.log('写入sheet文件失败')
    return;
  }
    // 发送邮件
    await sendEmail(subjectText);
  });
}, () => {
  // 定时任务回调函数
});
