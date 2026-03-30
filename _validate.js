const fs=require('fs');
const file=process.argv[2]||'template.html';
const html=fs.readFileSync(file,'utf-8');
const m=html.match(/<script>([\s\S]*?)<\/script>/gi);
if(!m){console.log('No scripts found');process.exit(1);}
m.forEach((s,i)=>{
  const code=s.replace(/<\/?script>/gi,'');
  if(code.trim().length<50)return;
  try{
    new Function(code);
    console.log('Script block '+(i+1)+': OK ('+code.length+' chars)');
  }catch(e){
    console.log('Script block '+(i+1)+': ERROR: '+e.message);
  }
});
