const CONFIG={APP_NAME:"Smart POS",SHEETS:{USERS:"Users",SETTINGS:"Settings",LOGS:"Logs",PRODUCTS:"Products"}};
const HEADERS={Users:["id","name","username","passwordHash","role","active","createdAt","updatedAt"],Settings:["key","value","updatedAt"],Logs:["id","userId","username","action","details","timestamp"],Products:["productId","productName","sku","category","unit","purchasePrice","salePrice","openingStock","minimumStock","status","createdAt","updatedAt"]};

function doGet(){setupSheets();return HtmlService.createHtmlOutputFromFile("Index").setTitle(CONFIG.APP_NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function setup(){setupSheets();seedAdmin();return "Step 2A setup completed.";}
function setupSheets(){
  const ss=SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(n=>{let s=ss.getSheetByName(n);if(!s)s=ss.insertSheet(n);if(s.getLastRow()===0)s.getRange(1,1,1,HEADERS[n].length).setValues([HEADERS[n]]);s.setFrozenRows(1);});
  if(!rows("Settings").some(x=>x.key==="appName"))append("Settings",{key:"appName",value:CONFIG.APP_NAME,updatedAt:new Date()});
}
function seedAdmin(){if(rows("Users").length)return;let d=new Date();append("Users",{id:id("USR"),name:"Administrator",username:"admin",passwordHash:hash("admin123"),role:"admin",active:true,createdAt:d,updatedAt:d});}
function login(username,password){
  setupSheets();username=String(username||"").trim();password=String(password||"");
  let u=rows("Users").find(x=>String(x.username).toLowerCase()===username.toLowerCase()&&String(x.active).toLowerCase()!=="false");
  if(!u||u.passwordHash!==hash(password))throw new Error("Invalid username or password.");
  let session={sessionId:Utilities.getUuid(),userId:u.id,name:u.name,username:u.username,role:u.role};
  CacheService.getScriptCache().put("POS_SESSION_"+session.sessionId,JSON.stringify(session),21600);log(session,"LOGIN","User logged in.");return session;
}
function logout(sid){let s=getSession(sid);if(s)log(s,"LOGOUT","User logged out.");if(sid)CacheService.getScriptCache().remove("POS_SESSION_"+sid);return true;}
function dashboard(sid){let s=requireSession(sid);return {appName:CONFIG.APP_NAME,user:{id:s.userId,name:s.name,username:s.username,role:s.role},serverTime:new Date().toString()};}

function getUsers(sid){let s=requireSession(sid);if(String(s.role).toLowerCase()!=="admin")throw new Error("Only administrator can access Users.");return rows("Users").map(x=>({id:String(x.id||""),name:String(x.name||""),username:String(x.username||""),role:String(x.role||""),active:x.active===true||String(x.active).toLowerCase()==="true",createdAt:x.createdAt?String(x.createdAt):""}));}
function saveUser(sid,d){let s=requireSession(sid);if(String(s.role).toLowerCase()!=="admin")throw new Error("Only administrator can create users.");let name=String(d.name||"").trim(),u=String(d.username||"").trim(),p=String(d.password||""),r=String(d.role||"staff");if(!name||!u||!p)throw new Error("Name, username and password are required.");if(p.length<6)throw new Error("Password must be at least 6 characters.");if(rows("Users").some(x=>String(x.username).toLowerCase()===u.toLowerCase()))throw new Error("Username already exists.");let n=new Date();append("Users",{id:id("USR"),name,username:u,passwordHash:hash(p),role:r,active:true,createdAt:n,updatedAt:n});log(s,"CREATE_USER","Created user: "+u);return true;}
function getSettings(sid){let s=requireSession(sid);if(String(s.role).toLowerCase()!=="admin")throw new Error("Only administrator can access Settings.");return rows("Settings").map(x=>({key:String(x.key||""),value:String(x.value??""),updatedAt:x.updatedAt?String(x.updatedAt):""}));}
function saveSetting(sid,key,value){let s=requireSession(sid);if(String(s.role).toLowerCase()!=="admin")throw new Error("Only administrator can change Settings.");key=String(key||"").trim();if(!key)throw new Error("Setting key is required.");let sh=SpreadsheetApp.getActive().getSheetByName("Settings"),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===key){sh.getRange(i+1,2).setValue(value);sh.getRange(i+1,3).setValue(new Date());log(s,"UPDATE_SETTING",key);return true;}append("Settings",{key,value,updatedAt:new Date()});log(s,"CREATE_SETTING",key);return true;}

/* STEP 2A — PRODUCT ADD */
function addProduct(sid,d){
  const s=requireSession(sid);
  const role=String(s.role||"").toLowerCase();
  if(role!=="admin"&&role!=="manager")throw new Error("You do not have permission to add products.");

  const productName=String(d.productName||"").trim();
  const sku=String(d.sku||"").trim();
  const category=String(d.category||"").trim();
  const unit=String(d.unit||"").trim();
  const purchasePrice=Number(d.purchasePrice);
  const salePrice=Number(d.salePrice);
  const openingStock=Number(d.openingStock);
  const minimumStock=Number(d.minimumStock);

  if(!productName)throw new Error("Product Name is required.");
  if(!sku)throw new Error("SKU / Barcode is required.");
  if(!category)throw new Error("Category is required.");
  if(!unit)throw new Error("Unit is required.");
  if(!Number.isFinite(purchasePrice)||purchasePrice<0)throw new Error("Enter a valid Purchase Price.");
  if(!Number.isFinite(salePrice)||salePrice<0)throw new Error("Enter a valid Sale Price.");
  if(!Number.isFinite(openingStock)||openingStock<0)throw new Error("Enter a valid Opening Stock.");
  if(!Number.isFinite(minimumStock)||minimumStock<0)throw new Error("Enter a valid Minimum Stock.");

  if(rows("Products").some(x=>String(x.sku).trim().toLowerCase()===sku.toLowerCase())){
    throw new Error("SKU / Barcode already exists.");
  }

  const now=new Date();
  const product={
    productId:id("PRD"),
    productName,
    sku,
    category,
    unit,
    purchasePrice,
    salePrice,
    openingStock,
    minimumStock,
    status:"Active",
    createdAt:now,
    updatedAt:now
  };

  append("Products",product);
  log(s,"CREATE_PRODUCT","Created product: "+productName+" | SKU: "+sku);

  return {success:true,product:{
    productId:product.productId,
    productName:product.productName,
    sku:product.sku,
    category:product.category,
    unit:product.unit,
    purchasePrice:product.purchasePrice,
    salePrice:product.salePrice,
    openingStock:product.openingStock,
    minimumStock:product.minimumStock,
    status:product.status
  }};
}

function getProducts(sid){
  requireSession(sid);
  return rows("Products").map(x=>({
    productId:String(x.productId||""),
    productName:String(x.productName||""),
    sku:String(x.sku||""),
    category:String(x.category||""),
    unit:String(x.unit||""),
    purchasePrice:Number(x.purchasePrice)||0,
    salePrice:Number(x.salePrice)||0,
    openingStock:Number(x.openingStock)||0,
    minimumStock:Number(x.minimumStock)||0,
    status:String(x.status||"Active")
  }));
}

function getSession(sid){if(!sid)return null;let r=CacheService.getScriptCache().get("POS_SESSION_"+sid);return r?JSON.parse(r):null;}
function requireSession(sid){let s=getSession(sid);if(!s)throw new Error("Session expired. Please login again.");return s;}
function log(u,a,d){append("Logs",{id:id("LOG"),userId:u.userId||u.id||"",username:u.username||"",action:a,details:d||"",timestamp:new Date()});}
function rows(n){let s=SpreadsheetApp.getActive().getSheetByName(n);if(!s||s.getLastRow()<2)return[];let h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0],v=s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues();return v.map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]])));}
function append(n,o){let s=SpreadsheetApp.getActive().getSheetByName(n);if(!s)throw new Error("Sheet not found: "+n);let h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];s.appendRow(h.map(x=>o[x]!==undefined?o[x]:""));}
function id(p){return p+"_"+Utilities.getUuid().replace(/-/g,"").slice(0,12).toUpperCase();}
function hash(p){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p),Utilities.Charset.UTF_8).map(b=>(b<0?b+256:b).toString(16).padStart(2,"0")).join("");}