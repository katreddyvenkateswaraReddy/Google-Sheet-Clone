// selecting DOM elements
const table = document.querySelector('#myTable');
const tableHeadingRow = document.querySelector('#table-heading-row');
const tbody = document.querySelector('#table-body');
const bold = document.querySelector('#bold');
const italic = document.querySelector('#italic');
const underline = document.querySelector('#underline');
const lineThrough = document.querySelector('#line-through');
const left = document.querySelector('#left');
const center = document.querySelector('#center');
const right = document.querySelector('#right');
const justify = document.querySelector('#justify');
const txtClr = document.querySelector('#txt-color');
const bgClr = document.querySelector('#bg');
const cut = document.querySelector('#cut');
const cpy = document.querySelector('#cpy');
const pst = document.querySelector('#pst');
const fontFamilies = document.querySelector('#family');
const f_size = document.querySelector('#size');
const cell = document.querySelector('.cell');
const home = document.querySelector('#home-menu');
const download = document.querySelector('#download');
const openFile = document.querySelector('#open');
const sheetBtnContainer = document.querySelector('.foot .sheets');

// matrix for virtual memory of sheet
const rows = 100;
const cols = 26;
let matrix = new Array(rows);


cell.innerText = 'A1';
let currentCell; // for current cell
let cutCpy = {}; // object to store data of cut or copied cell

// object for different font families
let fonts = {
    'Abril Fatface': 'cursive',
    'Amatic SC': 'cursive',
    'Cinzel': 'serif',
    'Cookie': 'cursive',
    'Crimson Text': 'serif',
    'Dancing Script': 'cursive',
    'Foldit': 'cursive',
    'Great Vibes': 'cursive',
    'Inconsolata': 'monospace',
    'Lato': 'sans-serif',
    'Lobster': 'cursive',
    'Major Mono Display': 'monospace',
    'Mogra': 'cursive',
    'Oswald': 'sans-serif',
    'Oxygen Mono': 'monospace',
    'Pacifico': 'cursive',
    'Playfair Display': 'serif',
    'Poppins': 'sans-serif',
    'Raleway': 'sans-serif',
    'Roboto': 'sans-serif',
    'Roboto Slab': 'serif',
    'Rubik': 'sans-serif',
    'Satisfy': 'cursive',
    'Shadows Into Light': 'cursive',
    'Smokum': 'cursive',
    'Times New Roman': 'serif',
    'Ubuntu Mono': 'monospace',
}
// font sizes
let fontSizes = [8, 10, 12, 14, 18, 20, 22, 24, 26, 30, 34, 38, 40, 42];

// generating matrix (virtual memory for excel data)
for (let i = 0; i < rows; i++) {
    matrix[i] = new Array(cols);
    for (let j = 0; j < cols; j++) {
        matrix[i][j] = {};
    }
}

// function for updating matrix on any changes
function updateMatrix(currentCell) {
    let obj = {
        style: currentCell.style.cssText,
        text: currentCell.innerText,
        id: currentCell.id,
    };
    let id = currentCell.id.split("");
    let i = id[1] - 1;
    let j = id[0].charCodeAt(0) - 65;
    matrix[i][j] = obj;
}

//function for downloading the excel data in json format
function downloadJson() {
    const jsonString = JSON.stringify(matrix);

    const blob = new Blob([jsonString], { type: 'application/json' }); // create file with the given data

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'data.json'; //giving a file name to the file that has to be downloaded

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

download.addEventListener('click', downloadJson);

// function for uploading json table data to excel and render it on webpage
//upload event
function openJson(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) { // reading the file user uploaded
            const fileContent = e.target.result;
            try {
                const data = JSON.parse(fileContent);

                matrix = data;
                data.forEach((row) => {
                    row.forEach((cell) => {
                        if (cell.id) {
                            let myCell = document.getElementById(cell.id)
                            myCell.innerText = cell.text;
                            myCell.style.cssText = cell.style
                        }
                    });
                })
            }
            catch (err) {
                console.log('Error in reading', err);
            }
        }
        reader.readAsText(file);
    }
}

openFile.addEventListener('change', openJson);

// variables required for generating and managing multiple sheet
let numSheets = 1;
let arrMatrix = [matrix];
let currSheetNum = 1;
let firstSheet = document.querySelector('.default'); // the 1st default sheet
let sheetBtn = [firstSheet]; //to get all the button created and manipulate styles
firstSheet.addEventListener('click', sheetListener)


// to save the current sheet into user local storage
document.querySelector('#save').addEventListener('click', ()=>{
    try{
        if (numSheets == 1) {
            let sheetArr = [matrix];
            localStorage.setItem('ArrMatrix', JSON.stringify(sheetArr));
        }
        else {
            let prevSheet = JSON.parse(localStorage.getItem('ArrMatrix'));
            let updated = [...prevSheet, matrix];
            localStorage.setItem('ArrMatrix', JSON.stringify(updated));
        }
        alert('Your current sheet is saved!!');
    }
    catch(err){
        alert('Could not save the sheet! \n', err)
    }
})

//  handling add new sheet to the page and rendering data
document.getElementById('add-btn').addEventListener('click', () => {
    
    // saving and storing sheets
    if (numSheets == 1) {
        let sheetArr = [matrix];
        localStorage.setItem('ArrMatrix', JSON.stringify(sheetArr));
    }
    else {
        let prevSheet = JSON.parse(localStorage.getItem('ArrMatrix'));
        let updated = [...prevSheet, matrix];
        localStorage.setItem('ArrMatrix', JSON.stringify(updated));
    }
    numSheets++;
    currSheetNum = numSheets;

    for (let i = 0; i < rows; i++) {
        matrix[i] = new Array(cols);
        for (let j = 0; j < cols; j++) {
            matrix[i][j] = {};
        }
    }

    // rendering new sheet
    tbody.innerHTML = '';

    for (let row = 1; row <= 100; row++) {
        let tr = document.createElement('tr');
        let th = document.createElement('th');
        th.innerText = row;
        th.style.backgroundColor = 'rgb(240, 240, 240)';
        tr.appendChild(th);
        for (let col = 1; col <= 26; col++) {
            let td = document.createElement('td');
            td.setAttribute('contenteditable', 'true');
            td.setAttribute('id', `${String.fromCharCode(col + 64)}${row}`);
            tr.appendChild(td);
            td.addEventListener('focus', (event) => { onfocus(event) });
            td.addEventListener("input", (event) => onInputFunction(event));
        }
        tbody.appendChild(tr);
    }
    
    // generating the sheet buttons
    let newSheet = document.createElement('button');
    newSheet.innerText = `Sheet ${currSheetNum}`;
    newSheet.setAttribute('id', currSheetNum-1)
    sheetBtnContainer.appendChild(newSheet);
    sheetBtn.push(newSheet);
    newSheet.addEventListener('click', sheetListener);
    sheetBtn.forEach((btn) => {
        if (btn.classList.contains('active-btn')) {
            btn.classList.remove('active-btn');
        }
    })
    newSheet.classList.add('active-btn');

})


// handling new sheet buttons generated dynamically while rendering new sheets
// rendering the target sheet on page
function sheetListener(event) {
    sheetBtn.forEach((btn) => {
        if (btn.classList.contains('active-btn')) {
            btn.classList.remove('active-btn');
        }
    })

    // to get data of only the sheet clicked
    tbody.innerHTML = '';

    for (let row = 1; row <= 100; row++) {
        let tr = document.createElement('tr');
        let th = document.createElement('th');
        th.innerText = row;
        th.style.backgroundColor = 'rgb(240, 240, 240)';
        tr.appendChild(th);
        for (let col = 1; col <= 26; col++) {
            let td = document.createElement('td');
            td.setAttribute('contenteditable', 'true');
            td.setAttribute('id', `${String.fromCharCode(col + 64)}${row}`);
            tr.appendChild(td);
            td.addEventListener('focus', (event) => { onfocus(event) });
            td.addEventListener("input", (event) => onInputFunction(event));
        }
        tbody.appendChild(tr);
    }
    
    event.target.classList.add('active-btn');
    let id = parseInt(event.target.id)
    console.log(id);

    // rendering the data of current sheet
    let myData = JSON.parse(localStorage.getItem('ArrMatrix'));
    let tableData = myData[id];
    matrix = tableData;
    if(!tableData){
        return;
    }
    tableData.forEach((row)=>{
        row.forEach((cell)=>{
            if(cell.id){
                let curCell = document.getElementById(cell.id);
                curCell.innerText = cell.text;
                curCell.style.cssText = cell.style;
            }
        });
    });
}

// generating font size options
fontSizes.forEach(size => {
    let opt = document.createElement('option');
    opt.value = size;
    opt.innerText = size;
    f_size.appendChild(opt);
})

// generating font family options
for (keys in fonts) {
    let opt = document.createElement('option');
    opt.value = `${keys}, ${fonts[keys]}`;
    opt.innerText = keys;
    opt.style.fontFamily = `${keys}, ${fonts[keys]}`;
    fontFamilies.appendChild(opt);
}

// generating the top heading A-Z
for (let i = 65; i <= 90; i++) {
    let th = document.createElement('th');
    th.innerText = String.fromCharCode(i);
    th.style.borderTop = 'none'
    tableHeadingRow.appendChild(th);
}


// generating all the excel rows
for (let row = 1; row <= 100; row++) {
    let tr = document.createElement('tr');
    let th = document.createElement('th');
    th.innerText = row;
    th.style.backgroundColor = 'rgb(240, 240, 240)';
    tr.appendChild(th);
    for (let col = 1; col <= 26; col++) {
        let td = document.createElement('td');
        td.setAttribute('contenteditable', 'true');
        td.setAttribute('id', `${String.fromCharCode(col + 64)}${row}`);
        tr.appendChild(td);
        td.addEventListener('focus', (event) => { onfocus(event) });
        td.addEventListener("input", (event) => onInputFunction(event));
    }
    tbody.appendChild(tr);
}

// making text bold
bold.addEventListener('click', () => {
    if (currentCell.style.fontWeight == 'bolder') {
        currentCell.style.fontWeight = 'normal'
        bold.classList.remove('active');
    }
    else {
        currentCell.style.fontWeight = 'bolder'
        bold.classList.add('active');
    }
    updateMatrix(currentCell);  
})

// making text italic
italic.addEventListener('click', () => {
    if (currentCell.style.fontStyle == 'italic') {
        currentCell.style.fontStyle = 'normal'
        italic.classList.remove('active');
    }
    else {
        currentCell.style.fontStyle = 'italic'
        italic.classList.add('active');
    }
    updateMatrix(currentCell);  
})

// underlining text
underline.addEventListener('click', () => {
    if (currentCell.style.textDecoration == 'underline') {
        currentCell.style.textDecoration = 'none'
        underline.classList.remove('active');
    }
    else {
        currentCell.style.textDecoration = 'underline'
        underline.classList.add('active');
    }
    updateMatrix(currentCell);  
})

// lign-through text 
lineThrough.addEventListener('click', () => {
    if (currentCell.style.textDecoration == 'line-through') {
        currentCell.style.textDecoration = 'none'
        lineThrough.classList.remove('active');
    }
    else {
        currentCell.style.textDecoration = 'line-through'
        lineThrough.classList.add('active');
    }
    updateMatrix(currentCell);  
})

// left alignment
left.addEventListener('click', () => {
    if (currentCell.style.textAlign == 'left') {
        left.classList.remove('active');
    }
    else {
        currentCell.style.textAlign = 'left'
        left.classList.add('active');
        right.classList.remove('active');
        center.classList.remove('active');
        justify.classList.remove('active');
    }
    updateMatrix(currentCell);  
})

// right alignment
right.addEventListener('click', () => {
    if (currentCell.style.textAlign == 'right') {
        currentCell.style.textAlign = 'left'
        right.classList.remove('active');
    }
    else {
        currentCell.style.textAlign = 'right'
        right.classList.add('active');
        left.classList.remove('active');
        center.classList.remove('active');
        justify.classList.remove('active');
    }
    updateMatrix(currentCell);  
})

// center alignment
center.addEventListener('click', () => {
    if (currentCell.style.textAlign == 'center') {
        currentCell.style.textAlign = 'left'
        center.classList.remove('active');
    }
    else {
        currentCell.style.textAlign = 'center'
        center.classList.add('active');
        left.classList.remove('active');
        right.classList.remove('active');
        justify.classList.remove('active');
    }
    updateMatrix(currentCell);  
})

// justify alignment
justify.addEventListener('click', () => {
    if (currentCell.style.textAlign == 'justify') {
        currentCell.style.textAlign = 'left'
        justify.classList.remove('active');
    }
    else {
        currentCell.style.textAlign = 'justify'
        justify.classList.add('active');
        left.classList.remove('active');
        center.classList.remove('active');
        right.classList.remove('active');
    }
    updateMatrix(currentCell);  
})

// current cell text color
txtClr.addEventListener('change', () => {
    currentCell.style.color = txtClr.value;
    document.querySelector('.textColoricn').style.borderColor = txtClr.value;
    updateMatrix(currentCell);  
})

// current cell background color 
bgClr.addEventListener('change', () => {
    currentCell.style.backgroundColor = bgClr.value;
    document.querySelector('.fillColoricn').style.borderColor = bgClr.value;
    updateMatrix(currentCell);  
})

// cut function
cut.addEventListener('click', () => {
    cutCpy = {
        style: currentCell.style.cssText,
        text: currentCell.innerText
    };
    currentCell.style = null;
    currentCell.innerText = '';
    updateMatrix(currentCell);  
})

// copy function
cpy.addEventListener('click', () => {
    cutCpy = {
        style: currentCell.style.cssText,
        text: currentCell.innerText
    };
})

// paste function
pst.addEventListener('click', () => {
    if (cutCpy.text) {
        currentCell.style = cutCpy.style;
        currentCell.innerText = cutCpy.text;
    }
    updateMatrix(currentCell);  
})

// changing font family
fontFamilies.addEventListener('change', () => {
    currentCell.style.fontFamily = fontFamilies.value;
    updateMatrix(currentCell);  
})

// changing font sizes
f_size.addEventListener('change', () => {
    currentCell.style.fontSize = `${f_size.value}px`;
    updateMatrix(currentCell);  
})

// the help button to display ui to send the query 
document.querySelector('#help-box-btn').addEventListener('click', ()=>{
    document.querySelector('.helpbox').style.display = 'flex';
})

//  to close the help box ui
document.querySelector('#close-btn').addEventListener('click', ()=>{
    document.querySelector('.helpbox').style.display = 'none';
})

// updating the matrix on each input
function onInputFunction(event) {
    updateMatrix(event.target);
}

// getting the focused cell and its id
function onfocus(e) {
    currentCell = e.target;
    cell.innerText = currentCell.id;
}
