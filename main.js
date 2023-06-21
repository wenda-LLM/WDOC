if (!document.location.href.endsWith("/index.html")) {
    throw new Error('非指定加载页面')
}
console.log("启动")
nw.App.clearCache()
document.title += "-OFFICE连接"
const http = require('http')
const server = http.createServer((req, res) => {
    const method = req.method
    const url = req.url
    console.log(method, url);

    if (url === '/' && method === 'GET') { // request请求命中路由1
        res.end('hello get!')
    } else if (url === '/' && method === 'POST') { // request请求命中路由2
        let bodyStr = ''
        req.on('data', (chunk) => {
            bodyStr += chunk
        })
        req.on('end', async () => {
            console.log(bodyStr)
            try {

                res.end(await rtst_客服(bodyStr.replace(/[\r\n]+/g, '\n')))
            } catch (error) {
                res.end(tooltip_window.app.content)
            }
            setTimeout(() => set_tooltip(""), 2000)

        })
    } else if (url === '/find' && method === 'POST') { // request请求命中路由2
        let bodyStr = ''
        req.on('data', (chunk) => {
            bodyStr += chunk
        })
        req.on('end', async () => {
            console.log(bodyStr)
            kownladge = (await find(bodyStr, app.zsk_step)).filter(i => !i.score || i.score < 170).map(i => ({
                title: get_title_form_md(i.title),
                url: get_url_form_md(i.title),
                content: i.content
            }))
            kownladge = kownladge.map((e, i) => i + 1 + ".《" + e.title + "》\n" + e.content).join('\n')
            res.end(kownladge)

        })
    } else if (url === '/raw' && method === 'POST') { // request请求命中路由2
        let bodyStr = ''
        req.on('data', (chunk) => {
            bodyStr += chunk
        })
        req.on('end', async () => {
            console.log(bodyStr)
            try {

                res.end(await send("raw!"+bodyStr))
            } catch (error) {
                res.end(tooltip_window.app.content)
            }
            setTimeout(() => set_tooltip(""), 2000)

        })
    } else { // request请求未命中以上路由
        res.end('404')
    }
    // 每次触发请求回调函数中只能调用一次response.end()，否则会报错
})
server.listen(17861, function () {
    console.log('Server is listening');
}) // 监听2000端口
gui = require('nw.gui')

gui.Window.open('tooltip.html', {
    "frame": false,
    "width": 400,
    "height": 800,
    "transparent": true,
    "resizable": false,
})
setTimeout(() => {
    gui.Window.getAll(l => window.tooltip_window = l[1].window)
}, 1000)
set_tooltip = s => {
    tooltip_window.app.content = s
}
window.global_onmessage = set_tooltip
nw.Window.get().on('close', () => nw.App.quit())