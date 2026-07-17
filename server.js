const express=require('express'),fs=require('fs'),path=require('path');
const multer=require('multer'),XLSX=require('xlsx');
const upload=multer({dest:path.join(__dirname,'data','uploads')});
const app=express(),PORT=3026,ADMIN_PWD='hangyou888';
app.use(express.json());
app.use(express.static(path.join(__dirname,'public')));

function rd(fn){try{return JSON.parse(fs.readFileSync(path.join(__dirname,'data',fn),'utf8'))}catch(e){return[]}}
function wd(fn,d){fs.writeFileSync(path.join(__dirname,'data',fn),JSON.stringify(d,null,2),'utf8')}
function auth(req,res,next){if(req.headers['x-admin-pwd']===ADMIN_PWD)return next();res.status(401).json({error:'需要管理员密码'})}

// 主页
app.get('/',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));

// 船员 CRUD
app.get('/api/crew',(req,res)=>res.json(rd('crew.json')));
app.put('/api/crew/:id',auth,(req,res)=>{
  let c=rd('crew.json');const i=c.findIndex(x=>x.id==req.params.id);
  if(i<0)return res.status(404).json({error:'未找到'});
  c[i]={...c[i],...req.body};wd('crew.json',c);res.json(c[i]);
});
app.post('/api/crew',auth,(req,res)=>{
  let c=rd('crew.json');const n={id:Date.now(),...req.body};c.push(n);wd('crew.json',c);res.json(n);
});
app.delete('/api/crew/:id',auth,(req,res)=>{
  let c=rd('crew.json');c=c.filter(x=>x.id!=req.params.id);wd('crew.json',c);res.json({ok:true});
});

// 操作日志
function crewLog(crewId,crewName,action,detail,oldVal,newVal){
  let log=rd('crew_log.json');
  log.push({id:Date.now(),crewId,crewName,action,detail,oldVal,newVal,time:new Date().toISOString()});
  wd('crew_log.json',log);
}

// 状态快捷变更
app.post('/api/crew/:id/status',auth,(req,res)=>{
  let c=rd('crew.json');const i=c.findIndex(x=>x.id==req.params.id);
  if(i<0)return res.status(404).json({error:'未找到'});
  const {status,ship,leaveDate}=req.body;
  const old=c[i];
  if(status){c[i].status=status;crewLog(c[i].id,c[i].name,'状态变更',old.status+' → '+status,old.status,status);}
  if(ship!==undefined){c[i].ship=ship;crewLog(c[i].id,c[i].name,'调船',old.ship+' → '+ship,old.ship,ship);}
  if(status==='已离职'&&leaveDate)c[i].leaveDate=leaveDate;
  if(status==='在船')c[i].joinDate=new Date().toISOString().split('T')[0];
  wd('crew.json',c);res.json(c[i]);
});

// 批量调船
app.post('/api/crew/batch-reassign',auth,(req,res)=>{
  const {crewIds,toShip}=req.body;
  if(!crewIds||!crewIds.length||!toShip)return res.status(400).json({error:'缺少参数'});
  let c=rd('crew.json');
  crewIds.forEach(cid=>{
    const i=c.findIndex(x=>x.id==cid);
    if(i>=0){crewLog(c[i].id,c[i].name,'调船',c[i].ship+' → '+toShip,c[i].ship,toShip);c[i].ship=toShip;}
  });
  wd('crew.json',c);res.json({ok:true,count:crewIds.length});
});

// 操作日志
app.get('/api/crew-log',(req,res)=>res.json(rd('crew_log.json')));

// 船舶 CRUD
app.get('/api/ships',(req,res)=>res.json(rd('ships.json')));
app.post('/api/ships',auth,(req,res)=>{
  let s=rd('ships.json');const n={id:Math.max(...s.map(x=>x.id),0)+1,...req.body};s.push(n);wd('ships.json',s);res.json(n);
});
app.put('/api/ships/:id',auth,(req,res)=>{
  let s=rd('ships.json');const i=s.findIndex(x=>x.id==req.params.id);
  if(i<0)return res.status(404).json({error:'未找到'});s[i]={...s[i],...req.body};wd('ships.json',s);res.json(s[i]);
});
app.delete('/api/ships/:id',auth,(req,res)=>{
  let s=rd('ships.json');s=s.filter(x=>x.id!=req.params.id);wd('ships.json',s);res.json({ok:true});
});

// 批量上传Excel/CSV更新船员
app.post('/api/crew/upload',auth,upload.single('file'),(req,res)=>{
  if(!req.file)return res.status(400).json({error:'请选择文件'});
  try{
    const ext=path.extname(req.file.originalname||'').toLowerCase();
    let wb;
    if(ext==='.csv'){
      const csvText=fs.readFileSync(req.file.path,'utf8');
      wb=XLSX.read(csvText,{type:'string'});
    }else{
      wb=XLSX.readFile(req.file.path);
    }
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    if(!rows||rows.length<2)return res.status(400).json({error:'文件为空或格式不正确'});

    // Header mapping: find column indices
    const header=rows[0].map(h=>String(h||'').trim());
    const idx={name:-1,age:-1,position:-1,ship:-1,phone:-1,status:-1,salary:-1,
      compNum:-1,compIssue:-1,compExpiry:-1,
      healthNum:-1,healthIssue:-1,healthExpiry:-1,
      trainNum:-1,trainIssue:-1,trainExpiry:-1,
      insType:-1,insNum:-1,insIssue:-1,insExpiry:-1};

    const map={
      '姓名':'name','船员姓名':'name','名字':'name',
      '年龄':'age',
      '职位':'position','职务':'position',
      '船舶':'ship','所属船舶':'ship','船名':'ship',
      '电话':'phone','手机':'phone','联系电话':'phone',
      '状态':'status','船员状态':'status',
      '薪资':'salary','月薪':'salary','工资':'salary',
      '适任证书编号':'compNum','适任证号':'compNum',
      '适任发证日期':'compIssue','适任签发日期':'compIssue',
      '适任到期日期':'compExpiry','适任有效期':'compExpiry',
      '健康证编号':'healthNum','健康证号':'healthNum',
      '健康证发证日期':'healthIssue','健康证签发日期':'healthIssue',
      '健康证到期日期':'healthExpiry','健康证有效期':'healthExpiry',
      '培训合格证编号':'trainNum','培训证号':'trainNum',
      '培训发证日期':'trainIssue','培训签发日期':'trainIssue',
      '培训到期日期':'trainExpiry','培训有效期':'trainExpiry',
      '保险险种':'insType','险种':'insType',
      '保险保单号':'insNum','保单号':'insNum',
      '保险生效日期':'insIssue','保险到期日期':'insExpiry',
    };

    header.forEach((h,i)=>{if(map[h]&&idx[map[h]]===-1)idx[map[h]]=i});

    if(idx.name===-1)return res.status(400).json({error:'未找到"姓名"列，请检查表头'});

    let crew=rd('crew.json');
    let created=0,updated=0,errors=[];

    for(let r=1;r<rows.length;r++){
      const row=rows[r];
      const name=String(row[idx.name]||'').trim();
      if(!name)continue;

      const status=idx.status>-1?String(row[idx.status]||'').trim():'';
      const validStatuses=['在船','休假','待派','已离职'];
      const finalStatus=validStatuses.includes(status)?status:'在船';

      const data={
        name,
        age:idx.age>-1?(parseInt(row[idx.age])||null):null,
        position:idx.position>-1?String(row[idx.position]||'').trim():'',
        ship:idx.ship>-1?String(row[idx.ship]||'').trim():'',
        phone:idx.phone>-1?String(row[idx.phone]||'').trim():'',
        status:finalStatus,
        salary:idx.salary>-1?(parseFloat(row[idx.salary])||null):null,
        certificates:{
          competency:{
            number:idx.compNum>-1?String(row[idx.compNum]||'').trim():'',
            issueDate:idx.compIssue>-1?String(row[idx.compIssue]||'').trim():'',
            expiryDate:idx.compExpiry>-1?String(row[idx.compExpiry]||'').trim():'',
          },
          health:{
            number:idx.healthNum>-1?String(row[idx.healthNum]||'').trim():'',
            issueDate:idx.healthIssue>-1?String(row[idx.healthIssue]||'').trim():'',
            expiryDate:idx.healthExpiry>-1?String(row[idx.healthExpiry]||'').trim():'',
          },
          training:{
            number:idx.trainNum>-1?String(row[idx.trainNum]||'').trim():'',
            issueDate:idx.trainIssue>-1?String(row[idx.trainIssue]||'').trim():'',
            expiryDate:idx.trainExpiry>-1?String(row[idx.trainExpiry]||'').trim():'',
          }
        },
        insurance:{
          type:idx.insType>-1?String(row[idx.insType]||'').trim():'',
          number:idx.insNum>-1?String(row[idx.insNum]||'').trim():'',
          issueDate:idx.insIssue>-1?String(row[idx.insIssue]||'').trim():'',
          expiryDate:idx.insExpiry>-1?String(row[idx.insExpiry]||'').trim():'',
        }
      };

      // Find existing crew by name
      const existIdx=crew.findIndex(c=>c.name===name);
      if(existIdx>=0){
        // Merge: only overwrite non-empty values
        const old=crew[existIdx];
        crew[existIdx]={...old,...data,
          certificates:{
            competency:{...old.certificates?.competency,...data.certificates.competency},
            health:{...old.certificates?.health,...data.certificates.health},
            training:{...old.certificates?.training,...data.certificates.training},
          },
          insurance:{...old.insurance,...data.insurance}
        };
        updated++;
      }else{
        data.id=Date.now()+Math.random();
        data.joinDate=new Date().toISOString().split('T')[0];
        crew.push(data);
        created++;
      }
    }

    wd('crew.json',crew);
    // Clean temp file
    fs.unlink(req.file.path,()=>{});
    res.json({ok:true,created,updated,errors,total:crew.length});
  }catch(e){
    console.error(e);
    try{fs.unlink(req.file.path,()=>{})}catch(_){}
    res.status(500).json({error:e.message});
  }
});

app.listen(PORT,'0.0.0.0',()=>{
  const os=require('os');const ifaces=os.networkInterfaces();let ip='';
  for(const name of Object.keys(ifaces)){for(const iface of ifaces[name]){if(iface.family==='IPv4'&&!iface.internal&&(name.includes('Wi')||name.includes('无线')||name.includes('WLAN')||!ip)){ip=iface.address}}}
  console.log('船员管理系统 '+PORT+'  http://'+(ip||'localhost')+':'+PORT);
});
