<div align="center">
    <img src="https://socialify.git.ci/imoki/wpsPush/image?description=1&font=Rokkitt&forks=1&issues=1&language=1&owner=1&pattern=Circuit%20Board&pulls=1&stargazers=1&theme=Dark">
<h1>金山推送器</h1>
基于「金山文档」的金山文档多渠道消息推送器

<div id="shield">

[![][github-stars-shield]][github-stars-link]
[![][github-forks-shield]][github-forks-link]
[![][github-issues-shield]][github-issues-link]
[![][github-contributors-shield]][github-contributors-link]

<!-- SHIELD GROUP -->
</div>
</div>

## 🍻 交流渠道  
<a href="https://space.bilibili.com/3546828310055281">B站：**无盐七**</a>  
QQ群：**963592267**  
公众号：**默库**  
  
## 🎊 简介
此脚本用于金山文档的消息推送，功能丰富，使用简单，推送类型繁多  

## ✨ 特性
    - 📀 支持金山文档运行
    - 💿 智能识别脚本类型、自动生成配置表格
    - ♾️ 多种推送方式、优先级推送、智能排版
    - 💽 使用简单、适配性强
    - 🔥 兼容airscript 1.0和airscript 2.0(Beta)

## 💬 支持得通知列表
- Bark（iOS）
- pushplus（微信）
- Server 酱（微信）
- 邮箱
- dingtalk（钉钉）
- discord
- 企业微信群机器人（企业微信）
- 息知（微信）
- 即时达（微信）
- wxpusher（微信）

## 📺️ 视频教程
[![](https://img.shields.io/badge/金山推送器-无盐七-blue)](https://www.bilibili.com/video/BV1bXckehEdn) https://www.bilibili.com/video/BV1bXckehEdn/
  
## 🛰️ 文字步骤
1. 复制最新UPDATE.js脚本到金山文档（脚本类型：airscript 1.0），并运行
2. 复制最新PUSH.js脚本到金山文档（脚本类型：airscript 1.0），添加网络API和邮箱API，并加入定时任务
3. 配置CONFIG表和PUSH表
4. 使用案例参考TEMPLATE.js脚本

## ⭐ 图片教程步骤
1. 复制UPDATE.js和PUSH.js到金山文档中。运行UPDATE脚本即可生成配置表格，给PUSH脚本添加网络API和邮箱API。  
![UPDATE脚本](https://s3.bmp.ovh/imgs/2024/07/16/b4d980655d4168f6.png)  
![PUSH脚本](https://s3.bmp.ovh/imgs/2024/07/16/36aa5198fb951403.png)  
![CONFIG表](https://s3.bmp.ovh/imgs/2024/07/16/a6ffc1e2c2ae10d9.png)  
2. PUSH表中“是否推送”选择“是”，CONFIG表中“推送方式”为“@all”会只推送填是的这几个  
![PUSH表](https://s3.bmp.ovh/imgs/2024/07/16/99e82f4b3e32f486.png)  
3. 将PUSH加入定时任务，到时间就会自动推送  
![定时任务](https://s3.bmp.ovh/imgs/2024/07/16/875218c387ce1dc0.png)  
4. 如何写入要推送的消息，参考TEMPLATE.js使用案例脚本即可  
![案例](https://s3.bmp.ovh/imgs/2024/07/16/3d164eeddd09d30b.png)  

## 🚀 推送逻辑流程
参考TEMPLATE.js使用案例脚本，将推送相关的代码复制到你的脚本中。  
当你的脚本调用**writeMessage**函数时，此函数会将消息写入CONFIG表中。  
等到PUSH定时任务执行时，会自动检索CONFIG表中的消息，并进行推送。  

## 🧾 表格配置含义
**任务的名称**  
writeMessage函数需要两个参数，taskName（任务名）和（message）消息  
此CONFIG表中的任务名称即为writeMessage需要的任务名称  
例如：CONFIG表中任务名称为“默山推送”  
那么使用时：writeMessage("待推送消息", "默山推送")   
  
**推送方式：**  
@all方式代表在PUSH表内的消息推送平台都推送（bark、pushplus、钉钉等等）  
bark方式代表，仅用推送bark  
dingtalk方式代表，仅用推送钉钉  
bark&pushplus方式，同时推送bark和pushplus。用&连接  
bark&email&pushplus方式，同时推送bark、email和pushplus。  
  
**加入消息池：**  
这个的意思是“加入消息池”选项勾选“是”的就会合并为一条消息进行通知，以@all方式推送。例如你运行了8个签到任务，那么在某个时刻只收到1条通知消息。  
默认为“否”，代表每个签到结果都用独立的一条消息通知。例如你运行了8个签到任务，那么在某个时刻会同时收到8条通知消息。  

## 🤝 欢迎参与贡献
欢迎各种形式的贡献

[![][pr-welcome-shield]][pr-welcome-link]

<!-- ### 💗 感谢我们的贡献者
[![][github-contrib-shield]][github-contrib-link] -->


## ✨ Star 数

[![][starchart-shield]][starchart-link]

## 📝 更新日志 
- 2025-01-12
    * 脚本类型同时兼容airscript 1.0和2.0版本
    * 推送脚本多功能更新
    * UPDATE.js脚本部分代码优化
- 2024-11-20
    * 增加wxpusher极简推送模式
    * 修复wxpusher不换行问题
    * 修复server酱不换行问题
- 2024-08-04
    * 增加wxpusher推送
- 2024-07-16
    * 推出金山文档多渠道消息推送器

<!-- ## 📌 特别声明

- 本仓库发布的脚本仅用于测试和学习研究，禁止用于商业用途，不能保证其合法性，准确性，完整性和有效性，请根据情况自行判断。

- 本人对任何脚本问题概不负责，包括但不限于由任何脚本错误导致的任何损失或损害。

- 间接使用脚本的任何用户，包括但不限于建立VPS或在某些行为违反国家/地区法律或相关法规的情况下进行传播, 本人对于由此引起的任何隐私泄漏或其他后果概不负责。

- 请勿将本仓库的任何内容用于商业或非法目的，否则后果自负。

- 如果任何单位或个人认为该项目的脚本可能涉嫌侵犯其权利，则应及时通知并提供身份证明，所有权证明，我们将在收到认证文件后删除相关脚本。

- 任何以任何方式查看此项目的人或直接或间接使用该项目的任何脚本的使用者都应仔细阅读此声明。本人保留随时更改或补充此免责声明的权利。一旦使用并复制了任何相关脚本或Script项目的规则，则视为您已接受此免责声明。

**您必须在下载后的24小时内从计算机或手机中完全删除以上内容**

> ***您使用或者复制了本仓库且本人制作的任何脚本，则视为 `已接受` 此声明，请仔细阅读*** -->

<!-- LINK GROUP -->
[github-codespace-link]: https://codespaces.new/imoki/wpsPush
[github-codespace-shield]: https://github.com/imoki/wpsPush/blob/main/images/codespaces.png?raw=true
[github-contributors-link]: https://github.com/imoki/wpsPush/graphs/contributors
[github-contributors-shield]: https://img.shields.io/github/contributors/imoki/wpsPush?color=c4f042&labelColor=black&style=flat-square
[github-forks-link]: https://github.com/imoki/wpsPush/network/members
[github-forks-shield]: https://img.shields.io/github/forks/imoki/wpsPush?color=8ae8ff&labelColor=black&style=flat-square
[github-issues-link]: https://github.com/imoki/wpsPush/issues
[github-issues-shield]: https://img.shields.io/github/issues/imoki/wpsPush?color=ff80eb&labelColor=black&style=flat-square
[github-stars-link]: https://github.com/imoki/wpsPush/stargazers
[github-stars-shield]: https://img.shields.io/github/stars/imoki/wpsPush?color=ffcb47&labelColor=black&style=flat-square
[github-releases-link]: https://github.com/imoki/wpsPush/releases
[github-releases-shield]: https://img.shields.io/github/v/release/imoki/wpsPush?labelColor=black&style=flat-square
[github-release-date-link]: https://github.com/imoki/wpsPush/releases
[github-release-date-shield]: https://img.shields.io/github/release-date/imoki/wpsPush?labelColor=black&style=flat-square
[pr-welcome-link]: https://github.com/imoki/wpsPush/pulls
[pr-welcome-shield]: https://img.shields.io/badge/🤯_pr_welcome-%E2%86%92-ffcb47?labelColor=black&style=for-the-badge
[github-contrib-link]: https://github.com/imoki/wpsPush/graphs/contributors
[github-contrib-shield]: https://contrib.rocks/image?repo=imoki%2Fsign_script
[docker-pull-shield]: https://img.shields.io/docker/pulls/imoki/wpsPush?labelColor=black&style=flat-square
[docker-pull-link]: https://hub.docker.com/repository/docker/imoki/wpsPush
[docker-size-shield]: https://img.shields.io/docker/image-size/imoki/wpsPush?labelColor=black&style=flat-square
[docker-size-link]: https://hub.docker.com/repository/docker/imoki/wpsPush
[docker-stars-shield]: https://img.shields.io/docker/stars/imoki/wpsPush?labelColor=black&style=flat-square
[docker-stars-link]: https://hub.docker.com/repository/docker/imoki/wpsPush
[starchart-shield]: https://api.star-history.com/svg?repos=imoki/wpsPython&type=Date
[starchart-link]: https://api.star-history.com/svg?repos=imoki/wpsPython&type=Date

