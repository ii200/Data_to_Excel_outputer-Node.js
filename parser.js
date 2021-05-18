'use strict'

const Excel             = require('exceljs')
const projectData       = require('./project_data')
const main_media_data   = require('./media_data')
const inner_media_data  = require('./inner_media')

const g_languages = ['en', 'es', 'it', 'pt', 'zh', 'fr', 'tr', 'de', 'ru', 'hu', 'pl']
// const g_languages = ['en'];

function GetPath(categoriesParsed) {
    for (let category in categoriesParsed.en) {
        let curr = category
        while (curr) {
            try {
                categoriesParsed.en[category].path = categoriesParsed.en[curr].path
                    ? [categoriesParsed.en[curr].category_placeholder, ...categoriesParsed.en[category].path]
                    : [categoriesParsed.en[curr].category_placeholder]
                if (categoriesParsed.en[curr].prev) {
                    if (categoriesParsed.en[categoriesParsed.en[curr].prev]) {
                        if (categoriesParsed.en[categoriesParsed.en[curr].prev].prev !== curr) {
                            curr = categoriesParsed.en[curr].prev;
                        } else {
                            console.log(curr, categoriesParsed.en[curr])
                            curr = null;
                        }
                    } else {
                        curr = null;
                    }
                }
                else {
                    curr = null
                }
            } catch (error) {
                console.log(error)
                console.log(curr, categoriesParsed.en[curr])
                curr = null

            }
        }
    }
    return categoriesParsed
}

function RmEmpty(categoriesParsed) {
    const bannedCateg = ['back', 'transformers', '9', '10', '9_1', '9_2', '9_3', '9_4', '10_1', '10_2', '10_3', '10_4']
    for (let categ in categoriesParsed["en"]) {
        const curr = categoriesParsed["en"][categ];
        if (!curr.category_placeholder) {
            for (let lang of g_languages) {
                delete categoriesParsed[lang][categ];
            }
        } else if (bannedCateg.includes(curr.category_placeholder.toLowerCase())) {
            // console.log('banned', curr.category_placeholder)
            for (let lang of g_languages) {
                delete categoriesParsed[lang][categ];
            }
        } 
    }

    return categoriesParsed
}

// TODO: rm second params
function RmSymbols(text, par) {
    try {
        return text.replace(/&reg;/g, '®').replace(/&trade;/g, '™').
        replace(/\<\/br\>/g, '\n').replace(/\<br\>/g, '\n').
        replace(/\<\/b\>/g, '').replace(/\<b\>/g, '').
        replace(/\<\/i\>/g, '').replace(/\<i\>/g, '').
        replace(/\<\/ul\>/g, '').replace(/\<ul\>/g, '').
        replace(/\<\/li\>/g, '').replace(/\<li\>/g, '\n • ').
        replace(/\<div\>/g, '\n').replace(/\<\/div\>/g, '\n')
    } catch (error) {
        // console.error(text, error)
        // console.log(par)
        return text
    }
}

function handleArray(arr) {
    if (!Array.isArray(arr)) return arr
    let res = '';
    for (let i = 0; i < arr.length; i++) {
        let word = arr[i];
        if (word && word !== true) {
            res += res ? (", " + word) : word;
        }
    }
    return res
}
// function BuildPath(categories) {
//     for (categ in categories["en"]) {
        
//     }
// }

function handleProp(prop, propName) {
    switch (propName) {
        case 'category_placeholder':
            return RmSymbols(handleArray(prop));
        case "title_page_text":
            return RmSymbols(handleArray(prop));
        case "automation_grey_m":
            return RmSymbols(handleArray(prop));
        case "automation_modal_window":
            return RmSymbols(handleArray(prop));
        case "modal_w_data_title":
            return RmSymbols(prop);
        case "modal_w_data_text":
            return RmSymbols(prop);
        case "modal_w_data_text":
            return RmSymbols(prop);
        case "video_text":
            return RmSymbols(prop);
        default: 
            return prop
                    
        
            


    }
}


function pParse() {
    let categories = {};
    //console.log(projectData)
    let structCategs = projectData[`endvr_data`]['categories'];
  //console.log(structCategs)
    categories["en"] = {
        "home": {},
        "smart_city_h": {
            prev: "home",
            category_placeholder: "Food and Beverage Factories",
            buttons: [],
        }
    };
    let smart_city_categs = ["dairy", "sugar_cane", "sugar_beet", "poultry", "grains","Beverage","beef","chocolate"]
    
    let industries_categs = [];

    for (let categ in structCategs) {
        categories["en"][categ] = categories["en"][categ] || {};
        
        // parent
        if (structCategs[categ].parent) {
            
            if (smart_city_categs.includes(categ.toLowerCase())) {
                categories["en"][categ].prev = "fb_ua";
            } else if (industries_categs.includes(categ.toLowerCase())) {
                categories["en"][categ].prev = "industry";
            } else {
                categories["en"][categ].prev = structCategs[categ].parent;
            }
        } else {
            console.log(categ, 'does not have a parent!')
        }
 
        // category_placeholder (header)
        if (structCategs[categ].category_placeholder) {
            categories["en"][categ]["category_placeholder"] = handleProp(structCategs[categ].category_placeholder, "category_placeholder");
        }

        // title page text
        if (structCategs[categ].title_page_text) {
            categories["en"][categ]["title_page_text"] = handleProp(structCategs[categ].title_page_text, "title_page_text");
        }

        // get english automation_grey_m
        if (structCategs[categ].automation_grey_m) {
                categories["en"][categ]["automation_grey_m"] = handleProp(structCategs[categ].automation_grey_m, "automation_grey_m");
        }

        if (structCategs[categ].automation_modal_window) {
            categories["en"][categ]["automation_modal_window"] = handleProp(structCategs[categ].automation_modal_window[0].small_md, "automation_modal_window");
        }

        /////// modal
        if (structCategs[categ].modal_w_data) {
            if (structCategs[categ].modal_w_data.title) {
                categories["en"][categ]["modal_w_data_title"] = handleProp(structCategs[categ].modal_w_data.title, "modal_w_data_title");
            }
            if (structCategs[categ].modal_w_data.overview_text) {
                categories["en"][categ]["modal_w_data_text"] = handleProp(structCategs[categ].modal_w_data.overview_text, "modal_w_data_text");
            }
        }

        if (structCategs[categ].languages) {
            for (let lang in structCategs[categ].languages) {
                const lang_ = lang.split('__')[1];
                categories[lang_] = categories[lang_] || {};
                categories[lang_][categ] = categories[lang_][categ] || {};
                for (let prop in structCategs[categ].languages[lang]) {
                    if (prop === "modal_w_data") {
                        if (structCategs[categ].languages[lang][prop]["title"]) {
                            categories[lang_][categ]["modal_w_data_title"] = handleProp(structCategs[categ].languages[lang][prop]["title"], "modal_w_data_title");
                        }
                        if (structCategs[categ].languages[lang][prop]["overview_text"]) 
                        {
                            categories[lang_][categ]["modal_w_data_text"] = handleProp(structCategs[categ].languages[lang][prop].overview_text, "modal_w_data_text");
                        } 
                        //
                    } 
                    else if (prop == "automation_modal_window") {
                        if(structCategs[categ].languages[lang][prop])
                        {
                            categories[lang_][categ]["automation_modal_window"] = handleProp(structCategs[categ].languages[lang][prop].small_md, "automation_modal_window");
                        }
                    } 
                    else {
                        categories[lang_][categ][prop] = handleProp(structCategs[categ].languages[lang][prop], prop);
                    }
                }
            }
        }

        if (structCategs[categ]['buttons']) {
            categories["en"][categ]["buttons"] = [];
            for (let buttonInd = 0; buttonInd <  structCategs[categ]['buttons'].length; buttonInd++) {
                
                let categ_ = categ;
                if (categ === "home") {
                    if  (smart_city_categs.includes(structCategs[categ]['buttons'][buttonInd][0].toLowerCase())) {
                        categ_ = "home";
                    } else if (industries_categs.includes(structCategs[categ]['buttons'][buttonInd][0].toLowerCase())) {
                        categ_ = "industry";
                    }
                }

                let button = structCategs[categ]['buttons'][buttonInd];

                let btnObj = {
                    categ: `${categ_}:${button[0]}`,
                    prev: categ_
                };

                
                if (button[5]) {
                    btnObj["button_title"] = button[5]; 
                } else if (button[4].button_options) {
                    if (button[4].button_options.title) {
                        btnObj["button_title"] = button[4].button_options.title;
                    }
                }
     
                if (button[4].carusel_text) {
                    btnObj["carusel_text"] = button[4].carusel_text;
                }

                if(button[4].video_text)
                {
                    btnObj["carusel_text"] = button[4].video_text.
                    replace(/\<\/br\>/g, '').replace(/\<br\>/g, '');
                }
                
                categories["en"][categ_]["buttons"][buttonInd] = btnObj;
                
                if (button[4].languages) {
                    for (let lang in button[4].languages) {
                        const lang_ = lang.split('__')[1];
                        categories[lang_] = categories[lang_] || {};
                        categories[lang_][categ_] = categories[lang_][categ_] || {};
                        categories[lang_][categ_]["buttons"] = categories[lang_][categ_]["buttons"] || [];
                        let btnObj_ = {
                            categ_: `${categ_}:${button[0]}`,
                            prev: categ_
                        };

                        for (let prop in button[4].languages[lang]) {
                            let propName = '';
                            if (prop === 'button_name__unbo') {
                                if (!button.button__title) {
                                    propName = 'button_title';
                                } else continue;
                            } else if (prop === 'button__title') {
                                propName = 'button_title';
                            } else {
                                propName = prop;
                            }
                            btnObj_[propName] = RmSymbols(handleArray(button[4].languages[lang][prop]));
                        }
                        categories[lang_][categ_]["buttons"][buttonInd] = btnObj_;
                        
                    }
                }
            }
        } 



    }
    return RmEmpty(GetPath(categories))
}

let categRes = pParse()
let workbook = new Excel.Workbook()
let worksheet = workbook.addWorksheet('Categories');
let cols = [];
cols = [
    // { header: 'categ', key: 'categ' },
    { header: 'Path', key: 'path' }
];
for (let lang of g_languages) {
    cols.push({ header: `Title ${lang}`, key: `title_page_text_${lang}` },
        { header: `Placeholder ${lang}`, key: `category_placeholder_${lang}` },
        { header: `Information circle ${lang}`, key: `automation_grey_m_${lang}` },
        { header: `Small box ${lang}`, key: `automation_modal_window_${lang}` },
        { header: `Title in modal window ${lang}`, key: `modal_w_data_title_${lang}` },
        { header: `Text in modal window ${lang}`, key: `modal_w_data_text_${lang}` },
        // { header: `Button placeholder ${lang}`, key: `button_name_unbo_${lang}` },
        { header: `Button placeholder ${lang}`, key: `button_title_${lang}` },
        { header: `Carusel and video text ${lang}`, key: `carusel_text_${lang}` },
        { header: `Links ${lang}`, key: `media_links_${lang}` },
        { header: `Media elements ${lang}`, key: `media_elements_${lang}` },
    );
}

worksheet.columns = cols;

let find_media_id = id =>
{
    for(let i = 0; i < inner_media_data["inner_media_data"].length; i++)
    {
        if(inner_media_data["inner_media_data"][i]["id"] == id)
        {
            return [inner_media_data["inner_media_data"][i],i];
        }
    }
}

// for (const categ in projectData[`endvr_data_en`]['categories']) {
// for (const categ in categRes.en) {

// for (const categ of Object.keys(categRes.en).sort()) { 
for (const categ in categRes.en) {
    if (categRes.en[categ]) {
        let rowData = {};
        cols.forEach(col => {
            if (col.key === 'path') {
                rowData["path"] = categRes.en[categ].path.join(' >> ');
            } else if (col.key === 'categ') {
                rowData[col.key] = categ;
            } else {
                let n = col.key.split('_').slice(0, -1).join('_');
                let lang = col.key.split('_').pop();
                //console.log(lang)
                if (categRes[lang][categ]) {
                    rowData[col.key] = categRes[lang][categ][n] || '';
                } else {
                    rowData[col.key] = '';
                }
            }
        });    
        worksheet.addRow(rowData);


        for (let buttonInd = 0; buttonInd < categRes.en[categ]["buttons"].length; buttonInd++) {
            let button = categRes.en[categ]["buttons"][buttonInd];
            if (!button) continue;
            let objData = {};
            cols.forEach(col => {
                if (col.key === 'path') {
                    objData["path"] = rowData.path + " >> " + button.button_title;
                } else if (col.key === 'categ') {
                    objData[col.key] = button.categ;
                } else if( col.key.indexOf("media_links_") > -1 || col.key.indexOf("media_elements_") > -1 ) {
                    let lang = col.key.split('_').pop();
                    //check media in category
                    {
                        if( projectData[`endvr_data`].categories[categ] )
                        if( projectData[`endvr_data`].categories[categ]["modal_w_data"] )
                        if( projectData[`endvr_data`].categories[categ]["modal_w_data"]["media_data"] )
                        {
                            let media_data = projectData[`endvr_data`].categories[categ]["modal_w_data"]["media_data"];
                            //
                            for(let k = 0; k < media_data.length; k++)
                            {
                                let id              = media_data[k];
                                let media_data_temp = find_media_id(id);
                                let media           = media_data_temp[0];
                                let tags            = eval(media.tags);
                                let available       = false;
                                
                                for(let y = 0; y < tags.length; y++)
                                {
                                    let ac_tag = tags[y];
                                    let out_lang = "";

                                    for(let u = 0; u < main_media_data.tags_data.length; u++)
                                    {
                                        if( ac_tag == main_media_data.tags_data[u].id )
                                        {
                                            out_lang = main_media_data.tags_data[u].name;
                                            break;
                                        }
                                    }

                                    if(out_lang)
                                    if(out_lang == "language" || out_lang == "all" || out_lang == lang)
                                    {
                                        available = true;
                                    }
                                }

                                if(available)
                                {
                                    if(col.key.indexOf("media_links_") > -1)
                                    {
                                        if(media.type == "link")
                                        {
                                            objData[col.key] = " • "+ media.url+'\n';
                                        }
                                    }
                                    else if(col.key.indexOf("media_elements_") > -1)
                                    {
                                        if(media.type != "link")
                                        {
                                            objData[col.key] = " • "+ media.file_name + " ("+ media.url +") " +'\n';
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else {
                    let n = col.key.split('_').slice(0, -1).join('_');
                    let lang = col.key.split('_').pop();
                    
                    if (categRes[lang][categ]) 
                    {
                        if (!categRes[lang][categ]["buttons"]) 
                        {
                            objData[col.key] = '';
                        } 
                        else if (categRes[lang][categ]["buttons"][buttonInd]) 
                        {
                            objData[col.key] = categRes[lang][categ]["buttons"][buttonInd][n] || '';
                        } 
                        else 
                        {
                            objData[col.key] = '';
                        }
                    }
                    else
                    {
                        objData[col.key] = '';
                    }
                }
                
            });
            worksheet.addRow(objData);

        }
        
    }
}

worksheet.columns.forEach(column => {
    column.width = 40
})

worksheet.getRow(1).font = { bold: true }
const insideColumns = {
    //  path: ['A'], 
    "en": ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', ],
    "es": ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',  ],
    "it": ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',  ],
    "pt": ['Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',  ],
    "zh": ['AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO',  ],
    "fr": ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW',  ],
    "tr": ['AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE',  ],
    "de": ['BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM',  ],
    "ru": ['BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU',  ],
    "hu": ['BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', ],
    "pl": ['CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', ]
};
worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    // TODO: change to H I when there is no categ column
    if (rowNumber !== 1 && (worksheet.getCell(`H${rowNumber}`).value || worksheet.getCell(`I${rowNumber}`).value)) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            cell.font = {
                color: { argb: 'ff5b5b5a' },
                italic: true
            } 
        })
    }

    for (let cols in insideColumns) {
        insideColumns[cols].forEach((v, i) => {
            let currCell = worksheet.getCell(`${v}${rowNumber}`);
            if (v === 'Q' && rowNumber === 25) 
            {
                console.log(currCell.value)
                console.log(worksheet.getCell(`${insideColumns["en"][i]}${rowNumber}`).value, `${insideColumns["en"][i]}${rowNumber}`)
            }
            currCell.border = {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'bold' },
                right: { style: 'bold' },
            }

            
            if (cols != 'en') { 
                if (typeof worksheet.getCell(`${insideColumns["en"][i]}${rowNumber}`).value === 'string' && 
                    worksheet.getCell(`${insideColumns["en"][i]}${rowNumber}`).value !== '') {
                        if (typeof currCell.value !== 'string' || currCell.value === '') {
                            currCell.fill = {
                                type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF0586a7'}
                            };
                        } else if (currCell.value === worksheet.getCell(`${insideColumns["en"][i]}${rowNumber}`).value) {

                            console.log('same', v, rowNumber)
                            currCell.fill = {
                                type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFc9f08d'}
                            };
                        }
                    }
            }
        })
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            cell.alignment = { wrapText: true };
        })
    }

})

workbook.xlsx.writeFile('Categ_parsed.xlsx')