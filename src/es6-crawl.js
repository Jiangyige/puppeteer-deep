const puppeteer = require('puppeteer');
var {timeout, moment} = require('../tools/tools.js');

const schedule = require('node-schedule');
var fs = require('fs');
var xlsx = require('node-xlsx');


// *  *  *  *  *  *
// ┬ ┬ ┬ ┬ ┬ ┬
// │ │ │ │ │  |
// │ │ │ │ │ └ day of week (0 - 7) (0 or 7 is Sun)
// │ │ │ │ └───── month (1 - 12)
// │ │ │ └────────── day of month (1 - 31)
// │ │ └─────────────── hour (0 - 23)
// │ └──────────────────── minute (0 - 59)
// └───────────────────────── second (0 - 59, OPTIONAL)

// 　　6个占位符从左到右分别代表：秒、分、时、日、月、周几

// 　　'*'表示通配符，匹配任意，当秒是'*'时，表示任意秒数都触发，其它类推

// 　　下面可以看看以下传入参数分别代表的意思

// 每分钟的第30秒触发： '30 * * * * *'

// 每小时的1分30秒触发 ：'30 1 * * * *'

// 每天的凌晨1点1分30秒触发 ：'30 1 1 * * *'

// 每月的1日1点1分30秒触发 ：'30 1 1 1 * *'

// 2016年的1月1日1点1分30秒触发 ：'30 1 1 1 2016 *'

// 每周1的1点1分30秒触发 ：'30 1 1 * * 1'


// let rule = new schedule.RecurrenceRule();

// // your timezone
// rule.tz = 'Asia/Shanghai';

// // runs at 15:00:00
// rule.second = 30;
// rule.minute = 0;
// rule.hour = 0;


let job = schedule.scheduleJob('0 0 9 * * *', () => {
  console.log('job schedule', new Date());

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

    console.log(powerArr);

    let powerNum = Number(powerArr.map(e => e.text).join(''))

    console.log(moment("Y-M-D h:m:s"), powerNum)

    await page.close()
    await browser.close();

    // 先查出历史数据
    var sheets = xlsx.parse('trd.xlsx');
      
    let beforeElsxData = []
    // 遍历 sheet
    sheets.forEach(function(sheet){
        // console.log(sheet['name']);
        // 读取每行内容
        for(var rowId in sheet['data']){
            // console.log(rowId);
            var row=sheet['data'][rowId];
            // console.log(row);
            beforeElsxData.push(row)
        }
    });

    beforeElsxData.push([
      moment("Y-M-D h:m:s"),
      powerNum,
    ])

    console.log(beforeElsxData);

    var data = [{
      name: 'sheet1',
      data: beforeElsxData
    }]
    var buffer = xlsx.build(data);
  
    // 写入文件
    fs.writeFile('trd.xlsx', buffer, function(err) {
      if (err) {
          console.log("Write failed: " + err);
          return;
      }
  
      console.log("Write completed.");


      const nodemailer = require('nodemailer');

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
          subject: 'TRD' + moment("Y-M-D") + '充电量', // 主题
          text: 'TRD' + moment("Y-M-D") + '充电量', // plain text body
          html: '<b>详情看附件</b>', // html body
          // 下面是发送附件，不需要就注释掉
          attachments: [{
                  filename: 'trd.xlsx',
                  path: './trd.xlsx'
              }
          ]
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


