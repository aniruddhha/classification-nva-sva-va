const XLSX = require('xlsx')

const nvaa = [
    { keywords: [`walk`], descrition: 'NVAA-Walk : Operator taking steps carrying no material' },
    { keywords: [`wait`], descrition: 'NVAA-Wait : Operator waiting due to some interfence or machine cycle' },
    { keywords: [`put`, `get`, `place`, `grasp`, `grab`], descrition: 'NVAA-Wait : Operator waiting due to some interfence or machine cycle' },
    { keywords: [`set`], descrition: 'NVAA-Handling - Set : Releasing item with no Value Added. Not Inserting/Assembling' },
    { keywords: [], descrition: 'NVAA-Handling - Double Handle : Releasing an item and picking it back up or switching hands' },
    { keywords: [`align`, `orient`], descrition: 'NVAA-Handling - Orient : Move material to align for another VA element, incl tent place' },
    { keywords: [`prepare`], descrition: 'NVAA-Handling - Prepping : Any element of work that gets part/item ready for VA element' },
    { keywords: [`toss`], descrition: 'NVAA-Handling - Trash : Handling or disposing of an item into trash' },
    { keywords: [`tool`], descrition: 'NVAA-Handling - Tool : Any elements of movement involved in obtaining and using tools' },
    { keywords: [`rework`], descrition: 'NVAA-Quality - Rework : Any elements of work to fix/repair an item/product' },
    { keywords: [`clean`], descrition: 'NVAA-Quality - Cleaning : Any elements of work to clean material/product' },
    { keywords: [`inspect`], descrition: 'NVAA-Quality - Inspect : Any elements of work checking for quality' },
    { keywords: [], descrition: 'NVAA-Off Standard : Off Standard' },
    { keywords: [], descrition: 'NVAA-Non Standard : Non Standard' },
]

const svaa = [
    { keywords: [`hold`], descrition: 'SVAA-Holding : Holding' },
    { keywords: [`pick`, `pt`, `get`, `place`, `grasp`, `grab`, `obtain`, `throw`], descrition: 'SVAA-Picking : Pick/Grab material considered SVAA if the material is presented in the AA Golden Zone / Strike Zone only' },
    { keywords: [`position`], descrition: 'SVAA-Positioning : Positioning' },
    { keywords: [`put`], descrition: 'SVAA-Putting Together : Putting Together' },
    { keywords: [`hammer`], descrition: 'SVAA-Hammering : Hammering' },
    { keywords: [`mask`, `unmask`], descrition: 'SVAA-Masking/Unmasking : Masking/Unmasking' },
    { keywords: [`grip`], descrition: 'SVAA-Gripping : Gripping' },
    { keywords: [`move`], descrition: 'SVAA-Moving (By Hand) : Moving (By Hand)' },
    { keywords: [`prepare`], descrition: 'SVAA-Preparing : Preparing' },
    { keywords: [`load`], descrition: 'SVAA-Loading : Loading' },
    { keywords: [`stretch out`], descrition: 'SVAA-Stretching Out : Stretching Out' },
    { keywords: [`unload`], descrition: 'SVAA-Unloading : Unloading' },
]

const vaa = [
    { keywords: [`screw`], descrition: 'VA-Secure/Fasten/Screw : Secure/Fasten/Screw' },
    { keywords: [`pressure`], descrition: 'VA-Applying Pressure : Applying Pressure' },
    { keywords: [`bend`], descrition: 'VA-Bending Material : Bending Material' },
    { keywords: [`weld`], descrition: 'VA-Welding : Welding' },
    { keywords: [`cut`], descrition: 'VA-Cutting Material : Cutting Material' },
    { keywords: [`glue`], descrition: 'VA-Glueing : Glueing' },
    { keywords: [`insert`], descrition: 'VA-Inserting : Inserting' },
    { keywords: [`install`, `assemble`], descrition: 'VA-Assemble/Sub Assemble : Assemble/Sub Assemble' },
    { keywords: [], descrition: 'VA-Production Identification : Production Identification' },
    { keywords: [], descrition: 'VA-Pretreatment : Pretreatment' },
    { keywords: [`paint`], descrition: 'VA-Painting : Painting' },
    { keywords: [`fit`], descrition: 'VA-Fitting : Fitting' },
    { keywords: [], descrition: 'VA-Quality Obligations (Legal) : Quality Obligations (Legal)' },
    { keywords: [`join`], descrition: 'VA-Joining : Joining' },
    { keywords: [`roll`], descrition: 'VA-Rolling : Rolling' },
]

const wb = XLSX.readFile(`classify.xlsx`)
const sheetRaw = wb.Sheets[wb.SheetNames[0]]
const jsonSheet = XLSX.utils.sheet_to_json(sheetRaw)

jsonSheet.map(row => {
    if (row.Category.toLowerCase() == 'nva') {
        nvaa.forEach(chk => {
            chk.keywords.forEach(kw => {
                if (row.Description.toLowerCase().includes(kw.toLowerCase())) {
                    row.type = chk.descrition
                    return
                }
            })
        })
    } else if (row.Category.toLowerCase() == 'sva') {
        svaa.forEach(chk => {
            chk.keywords.forEach(kw => {
                if (row.Description.toLowerCase().includes(kw.toLowerCase())) {
                    row.type = chk.descrition
                    return
                }
            })
        })
    } else {
        vaa.forEach(chk => {
            chk.keywords.forEach(kw => {
                if (row.Description.toLowerCase().includes(kw.toLowerCase())) {
                    row.type = chk.descrition
                    return
                }
            })
        })
    }
})

console.log(jsonSheet)
const wbNw = XLSX.utils.book_new()
const updates = XLSX.utils.json_to_sheet(jsonSheet)

XLSX.utils.book_append_sheet(wbNw, updates, 'transformed')
XLSX.writeFile(wbNw, 'book.xlsx')


