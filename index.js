function splitColum(props) {
 let c = {};
 for(let i in props){
  c[i]= props[i].match(/[A-Z]/gi).join('');
 }
 return c;
}

function splitRow(props) {
 let r = {};
 for(let i in props){
  r[i]= props[i].match(/\d/g).join('');
 }
 return r;
}

class ExcelScript {

 constructor(file, DataShop, DataProduct){
    this.xlsx,
    this.posts = [],
    this.post = {};
    this.fileName = file;
    if(Object(DataShop).length !== 'undefined'){
     this.DataShop = {
      ...DataShop
     }
    }
    if(Object(DataShop).length !== 'undefined'){
     this.DataProduct = {
      ...DataProduct
     }
    }
    this.upCase(this.DataShop, this.DataProduct)
    this.worksheet
    this.reqParse(this.fileName)
    if(Object(DataShop).length !== 'undefined'){
     this.method(this.worksheet, this.post)
    }
    if(Object(DataProduct).length !== 'undefined'){
     this.productMethod(this.worksheet, this.DataProduct)
    }
 
 }
 upCase(DataShop, DataProduct){
  if(Object(DataShop).length !== 'undefined'){
   for(let i in DataShop){
    this.DataShop[i] = DataShop[i].toUpperCase();
   }
  }
  if(Object(DataProduct).length !== 'undefined'){
   for(let i in DataProduct){
    this.DataProduct[i] = DataProduct[i].toUpperCase();
   }
  }
 }
 reqParse(file){
  this.xlsx = require('xlsx');
  const workbook = this.xlsx.readFile(file);
  this.worksheet = workbook.Sheets[workbook.SheetNames[0]];
 }
 method(worksheet, post){
  worksheet['!merges'].map((d, i)=>{
 let deCell = {};
 deCell = d.s;
 //deReng = this.xlsx.utils.encode_range(d);
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.nameShop){
   post.name = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.adressShop) {
   post.adress = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.InnShop){
   post.inn = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.KppShop){
   post.kpp = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.Shipper){
   post.shipper = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.deliveryPoint){
   post.delivery = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }
  if(this.xlsx.utils.encode_cell(deCell) === this.DataShop.checkingAccount){
   post.check = worksheet[this.xlsx.utils.encode_cell(deCell)].v;
  }

})
this.posts.push(post);
post = {};
}
productMethod(worksheet, DataProduct){
 let post = {};
 let data = this.xlsx.utils.sheet_to_json(worksheet, {range: `${DataProduct.start}:${DataProduct.end}`, header:'A', blankrows:false,});
  let column = {
   ...splitColum(DataProduct)
  }
for(let cell in data){

 for(let newCell in data[cell]){

    if(newCell === column.id) {
     post.ID = data[cell][newCell]
    }
    if(newCell === column.nameProduct) {
     post.nameProduct = data[cell][newCell];
    }
    if(newCell === column.quantity) {
     post.quantity = data[cell][newCell];
    }
    if(newCell === column.price) {
     post.price = data[cell][newCell];
    }
    if(newCell === column.totalPice) {
     post.totalPice = data[cell][newCell];
    }
    if(newCell === column.barCode) {
     post.barCode = data[cell][newCell];
    }
    
 }
 if(Object.keys(post).length !== 0){
  this.posts.push(post);
  post = {};
 }
}

}
}

const excelscript = new ExcelScript(
 './test.xlsx',
 DataShop = {
 nameShop:'h1', // наименование поставщика 
 adressShop:'H2', // адрес
 InnShop:'h3', // ИНН
 KppShop:'', // КПП
 Shipper:'h4', // пукт отправки груза
 deliveryPoint:'h5', // адрес доставки
 checkingAccount: ''  // расчетный счет
},
DataProduct = {
start:'a12', // Начальная координата таблицы
end: 'BT31', // Конечная координата таблицы
id:'C12', // id товара или артикул
nameProduct:'h12', // название продукта(наименование)
quantity:'AG12',// количество продукта в остатке
price:'AT12', // цена продукта
totalPice:'BH12', // суммарная стоимость продукта с НДС
barCode:'BT12' // штрихкод
}
)

const y = excelscript.posts;
console.log('y: ', y);



// Вариант решения №2 для себя============================================
 //for(let cell in worksheet){
 // const cellAsString = cell.toString();
 // const c = cellAsString.match(/[A-Z]/gi).join('');
 // let r;
 // let column = {
   //...splitColum(DataProduct)
  //}
  //let row = {
   //...splitRow(DataProduct)
  //}
   //if(cellAsString[1] !== 'r' && cellAsString[1] !== 'm'){
   // r = cellAsString.match(/\d/g).join('');
   //}
   
   //if(cellAsString[1] !== 'r' && cellAsString !== 'm' && +r >= +row.id && +r <= 22){

    //if(c === column.id) {
    // post.ID = worksheet[cell].v;
    //}
    //if(c === column.nameProduct) {
    // post.nameProduct = worksheet[cell].v;
    //}
    //if(c === column.quantity) {
    // post.quantity = worksheet[cell].v;
    //}
    //if(c === column.price) {
    // post.price = worksheet[cell].v;
    //}
    //if(c === column.totalPice) {
    // post.totalPice = worksheet[cell].v;
    //}
    //if(c === column.barCode) {
    // post.barCode = worksheet[cell].v;
    //}

    //this.posts.push(post);
     //post = {};
   //}
  
 //}
 

 //console.log(this.posts);

 // Вариант решения №2============================================
