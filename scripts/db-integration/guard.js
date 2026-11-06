const assert=require('node:assert/strict');
function guard(config){
 assert.equal(config.isolated,true);assert.equal(config.database,'humbel_isolated_test');assert.equal(config.host,'127.0.0.1');assert.equal(config.port,55432);
 for(const key of['auth','rest','edge','worker']){const u=new URL(config[key]);assert.equal(u.protocol,'http:');assert.equal(u.hostname,'127.0.0.1');assert.equal(u.username,'');assert.equal(u.password,'');}
 assert.ok(typeof config.password==='string'&&config.password.length>=48);assert.ok(typeof config.jwtSecret==='string'&&config.jwtSecret.length>=48);assert.match(config.sentinel,/^[a-f0-9]{64}$/);
 assert.ok(!JSON.stringify(config).includes('itntnpzqzyjfdnjdshvm'));return config;
}
async function verifyTarget(db,config){guard(config);const r=await db.query("select current_database() name,current_setting('humbel.isolated_sentinel',true) sentinel");assert.equal(r.rows[0].name,config.database);assert.equal(r.rows[0].sentinel,config.sentinel);}
module.exports={guard,verifyTarget};
