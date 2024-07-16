/*
    作者: imoki
    仓库: https://github.com/imoki/wpsPush
    公众号：默库
    更新时间：20240716
    脚本：使用案例参考
    说明：此脚本为使用案例，将推送相关的代码复制到你的脚本中
          然后调用writeMessage函数即可使用
    其他：关注默库官方渠道即时获取最新更新消息
*/


// 使用：
// 只需要向填写两个参数即可，taskName（任务名）和（message）消息
// 之后运行PUSH脚本就会自动进行推送了
let taskName = "推送任务1"  // 填CONFIG表的任务名，代表向CONFIG表中的次任务写入
let message = "这是一条消息"  // 填写待推送的消息
writeMessage(message, taskName)  // 将消息写入CONFIG表中


// 将如下内容复制到你的脚本中即可调用
// =================推送相关开始===================
// 获取时间
function getDate(){
  let currentDate = new Date();
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  return currentDate
}

// 将消息写入CONFIG表中作为消息队列，之后统一发送
function writeMessage(message, taskName){
  // 当天时间
  let todayDate = getDate()
  let sheetNameConfig = "CONFIG"; // 总配置表
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("✨ 开始将消息结果写入主配置表");
    for (let i = 0; i <= 100; i++) {  // 限制CONFIG为100行以内
      // 找到指定的表行
      if(Application.Range("A" + (i + 2)).Value == taskName){
        // 写入更新的时间
        Application.Range("C" + (i + 2)).Value = todayDate
        // 写入消息
        Application.Range("D" + (i + 2)).Value = message
        console.log("✨ 写入消息结果完成");
        break;  // 找到就提前退出
      }

      if(Application.Range("A" + (i + 2)).Value == ""){
        break;  // 空行提前退出，提高效率
      }
    }
  }
}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}
// =================推送相关结束===================