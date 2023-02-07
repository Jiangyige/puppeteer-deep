const puppeteer = require('puppeteer');
const {timeout, moment} = require('../tools/tools.js');

const schedule = require('node-schedule');
const fs = require('fs');
const xlsx = require('node-xlsx');
const nodemailer = require('nodemailer');


const rule = new schedule.RecurrenceRule();
// runs at 00:00:00
// rule.hour = 0;
// rule.minute = 0;
rule.second = 0;
rule.tz = 'Asia/Shanghai';

let job = schedule.scheduleJob(rule, () => {
  console.log('job schedule', new Date().toLocaleString());

  puppeteer.launch({args: ['--no-sandbox', '--disable-setuid-sandbox']}).then(async browser => {
    let page = await browser.newPage();

    await page.goto('https://www.teld.cn');
    await timeout(3000);

    console.log('start!');
    let powerArr = await page.evaluate(() => {
      let as = [...document.querySelectorAll('#Power p.number')];

      return as.map((a) =>{
          return {
            text: a.textContent
          }
      });
    });

    let powerNum = Number(powerArr.map(e => e.text).join(''))
    console.log(moment("Y-M-D h:m:s"), powerNum)

    await page.close()
    await browser.close();

    // 先查出历史数据
    const sheets = xlsx.parse('trd.xlsx');
      
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
    
    const subjectText = `时间：${moment("Y-M-D h:m:s")}, 充电量：${powerNum}, 增量：${powerNum - lastRowData[1]}`

    console.log(subjectText)

    beforeElsxData.push([
      moment("Y-M-D h:m:s"),
      powerNum,
      powerNum - lastRowData[1]
    ])

    const data = [{
      name: 'sheet1',
      data: beforeElsxData
    }]
    const buffer = xlsx.build(data);
  
    // 写入文件
    fs.writeFile('trd.xlsx', buffer, function(err) {
      if (err) {
          console.log("Write failed: " + err);
          return;
      }
  
      console.log("Write completed.");

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
          text: 'TRD' + moment("Y-M-D") + '充电量', // plain text body
          html: '<b>详情看附件</b>', // html body
          // 下面是发送附件，不需要就注释掉
          attachments: [{
            filename: 'trd.xlsx',
            path: './trd.xlsx'
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
    });
  });
});


