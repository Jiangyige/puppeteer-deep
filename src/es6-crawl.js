const puppeteer = require('puppeteer');
var {timeout, moment} = require('../tools/tools.js');

const schedule = require('node-schedule');

// 当前时间的秒值为 10 时执行任务，如：2021-7-8 13:25:10
let job = schedule.scheduleJob('30 * * * * *', () => {
  console.log('job schedule', new Date());

  puppeteer.launch().then(async browser => {
    let page = await browser.newPage();

    await page.goto('https://www.teld.cn');
    await timeout(3000);

    console.log('start111');
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
  });

});


