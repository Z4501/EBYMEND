let orders = []
let shipping = {}
let vendor = {}

function parseCSV(text){

let rows = text.split("\n")
let headers = rows[0].split(",")

let data=[]

for(let i=1;i<rows.length;i++){

let cols = rows[i].split(",")

if(cols.length<headers.length) continue

let obj={}

headers.forEach((h,index)=>{
obj[h.trim()] = cols[index]
})

data.push(obj)

}

return data
}

function processFiles(){

let ordersFile = document.getElementById("ordersFile").files[0]
let shipFile = document.getElementById("shipFile").files[0]

if(!ordersFile){
alert("Upload orders CSV")
return
}

let reader = new FileReader()

reader.onload = function(e){

orders = parseCSV(e.target.result)

loadShipping(shipFile)

}

reader.readAsText(ordersFile)

}

function loadShipping(file){

if(!file){
buildTable()
return
}

let reader = new FileReader()

reader.onload=function(e){

let data = parseCSV(e.target.result)

data.forEach(row=>{
shipping[row["Order #"]] = parseFloat(row["Shipping Cost"]||0)
})

buildTable()

}

reader.readAsText(file)

}

function estimateFee(price){

return price*0.13
}

function buildTable(){

let tbody = document.querySelector("#resultTable tbody")

tbody.innerHTML=""

let totalProfit=0
let totalSales=0

orders.forEach(o=>{

let order=o["Order Number"] || o["Order number"]

let sale=parseFloat(o["Sold For"]||0)

let ship=parseFloat(o["Shipping And Handling"]||0)

let tax=parseFloat(o["eBay Collected Tax"]||0)

let sku=o["Custom Label"]||""

let vendorCost=0

let shipCost=shipping[order]||0

let fee=estimateFee(sale)

let profit = sale + ship - fee - vendorCost - shipCost

let margin = (profit/sale)*100

totalProfit+=profit
totalSales+=sale

let tr=document.createElement("tr")

tr.innerHTML=`
<td>${order}</td>
<td>${sku}</td>
<td>${sale.toFixed(2)}</td>
<td>${ship.toFixed(2)}</td>
<td>${tax.toFixed(2)}</td>
<td>${vendorCost.toFixed(2)}</td>
<td>${shipCost.toFixed(2)}</td>
<td>${fee.toFixed(2)}</td>
<td class="${profit>0?'good':'bad'}">${profit.toFixed(2)}</td>
<td>${margin.toFixed(1)}%</td>
`

tbody.appendChild(tr)

})

document.getElementById("summary").innerHTML=`
Total Sales: $${totalSales.toFixed(2)}<br>
Total Profit: $${totalProfit.toFixed(2)}
`

}
