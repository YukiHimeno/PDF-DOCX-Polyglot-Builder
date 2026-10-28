# PDF-DOCX Polyglot Builder

纯浏览器端 PDF / DOCX 双格式文件构造器。文件不会上传到服务器，两个下载按钮导出同一个 Blob，并展示其 SHA-256。 产物文件改后缀名就可以作为 PDF / DOCX 打开。 

## 运行

需要 Node.js 20.19+ 或 22.12+。

```sh
npm install
npm run dev
npm test
npm run build
```

DOCX 自动转换需要跨源隔离（COOP/COEP 响应头），请注意检查部署配置。