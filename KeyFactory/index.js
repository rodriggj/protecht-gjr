const axios = require('axios')
const reader = require('xlsx')
const colors = require('colors')

// Global Variables -- Replace for each Affliate you intend to Onboard Clients too
const affl_info = {
    affiliate_id: 'af_Ri7RfjWt',
    public_key: 'pk_sandbox_ca05513a308341b07ec208a51d698bb268c1b5a0', 
    secret_key: 'sk_sandbox_ea060ec76ed971b2afded7ea521e84dd6e34db5a',
    path: `../Data/testData7.xlsx`
}

// -----------------------------------------------------------------------

// Function for retrieving the contents of an Excel document (extension .xlsx) 
const getFileContents = (filePath) => {

    // Reading our test file
    const file = reader.readFile(filePath)

    // Create a variable to hold information read from Excel File
    let data = []

    // Create an array containing the Sheet names contained in the file
    const sheets = file.SheetNames

    // Utilize the Utils function in the XlSX package to convert the contents of the excel file to JSON
    for(let i = 0; i < sheets.length; i++) {
        const temp = reader.utils.sheet_to_json(
            file.Sheets[file.SheetNames[i]])
            temp.forEach((res) => {
            data.push(res)
        })
        // console.log(data)
        return data
    }
}

// Function for retrieving an Authentication token from the Protecht v2 Authorization API
const getAffiliateToken = async (affl_pk, affl_sk) => {
    let config = {
        method: 'POST', 
        url: 'https://connect-sandbox.ticketguardian.net/api/v2/auth/token',
        headers: {
            "accept": 'application/json',
            "content-type": 'application/json'
        }, 
        data: {
            public_key: affl_pk, 
            secret_key: affl_sk
        }
    }
    const res = await axios(config)
    return res.data.token
}

// Function for retrieving a Client Public Key and Secret Key from the Protecht v2 Onboarding API
const getClientKeys = async (affl_id, affl_token, client) => {
    let config = {
        method: 'POST', 
        url: "https://connect-sandbox.ticketguardian.net/api/v2/accounts/",
        headers: {
            "authorization": `Bearer ${affl_token}`,
            "cache-control": `no-cache`,
            "content-type": "application/json"
        },
        data: {
            name: client.Client, 
            domain: client.Domain, 
            affiliate: affl_id,
            user: {
                email: client.Email,
                password: 'protecht!123',
            },
            product: client.Product,
            send_activation: true
        }
    } 
    const response = await axios(config);
    return response.data.api_keys
}

async function main() {
    try{
        // Get file contents returns as an Array
        const data = await getFileContents(affl_info.path)
        console.log(`******** 1. Data read from file ***********`.cyan)
        //console.log(data)

        // Create Array Data Structure for API response for Public and Secret Keys and initialize to empty
        let newData = []

        // Get Affiliate Authentication Token which is needed for Client Onboarding API calls -- need to add a time out of 3 minutes to the resolve of this Promise
        const auth_token = await getAffiliateToken(affl_info.public_key, affl_info.secret_key)

        console.log(`******** 2. Affiliate Authentication Token Retrieved ***********`.cyan)
        console.log(auth_token)

        // Create for loop to iterate through the elements of the file content (aka data variable), and make a request for Public & Secret keys
        console.log(`*********     3. New Records written to the workbook...   ********* `.cyan)
        for(let x = 0; x<data.length; x++) {
            const keys = await getClientKeys(affl_info.affiliate_id, auth_token, data[x])   // Loop initiates a API call to Onboarding API
            const { public_key, secret_key } = keys   // Add new Keys to existing Client Object
            let newClient = {...data[x], public_key, secret_key}   // Create new Client Object
            newData.push(newClient)   // push new Client to array
            console.log(`Client ${data[x].Client} issued public key: ${newData[x].public_key} and secret key: ${newData[x].secret_key}`.magenta)   //User is informed of API status on console
            setTimeout(()=>{console.log(`...`), 10000})   //Timeout is initiated for 10 miliseconds prior to the next API call
        }

        // Write the New Data back to the File (https://javascript.plainenglish.io/read-write-excel-file-in-node-js-using-xlsx-ab11881d00b4)
        const ws = reader.utils.json_to_sheet(newData)
        const wb = reader.utils.book_new()
        reader.utils.book_append_sheet(wb, ws, "Sheet1")
        reader.writeFile(wb, affl_info.path)
        console.log(`*********     4. New Records written to the workbook...   ********* `.cyan)
    } catch(error) {
        console.log(`${error}` )
    }
}

main()