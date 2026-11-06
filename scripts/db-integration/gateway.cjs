// Synthetic internal-only router: no arbitrary upstream, no credentials in code/logs.
const http=require('node:http');
http.createServer((req,res)=>{const prefix=req.url.startsWith('/auth/v1/')?'/auth/v1':req.url.startsWith('/rest/v1/')?'/rest/v1':null;if(!prefix){res.writeHead(404).end();return;}const upstream=http.request({hostname:prefix==='/auth/v1'?'humbel-auth':'humbel-rest',port:prefix==='/auth/v1'?9999:3000,path:req.url.slice(prefix.length),method:req.method,headers:req.headers},r=>{res.writeHead(r.statusCode,r.headers);r.pipe(res);});upstream.on('error',()=>res.writeHead(502).end());req.pipe(upstream);}).listen(8080,'0.0.0.0');
// Tests execute inside this container. No ports are published to the host or external networks.
const net=require('node:net');
for(const [listen,host,port]of[[55432,'humbel-db',5432],[59999,'humbel-auth',9999],[53000,'humbel-rest',3000],[59000,'humbel-edge',9000],[59001,'humbel-worker',9000]])net.createServer(socket=>{const upstream=net.connect(port,host);socket.on('error',()=>upstream.destroy());upstream.on('error',()=>socket.destroy());socket.pipe(upstream).pipe(socket);}).listen(listen,'127.0.0.1');
