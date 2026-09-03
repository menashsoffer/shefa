const http=require('http'),fs=require('fs'),p=require('path');
const root=__dirname;
http.createServer((req,res)=>{
  let f=decodeURIComponent(req.url.split('?')[0]); if(f==='/')f='/index.html';
  const fp=p.join(root,f);
  fs.readFile(fp,(e,d)=>{
    if(e){res.writeHead(404);res.end('nf');return;}
    const ext=p.extname(fp);
    const ct={'.html':'text/html; charset=utf-8','.json':'application/json','.js':'text/javascript','.css':'text/css','.geojson':'application/json'}[ext]||'application/octet-stream';
    res.writeHead(200,{'content-type':ct});res.end(d);
  });
}).listen(8777,()=>console.log('http://localhost:8777'));
