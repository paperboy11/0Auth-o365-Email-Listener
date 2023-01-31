CODE BY https://gazellia.com/  

module.exports = function (Config, Modules) {

    const msal = require('@azure/msal-node'); 
    var EmailListener = {};

    EmailListener.getAuthToken = function(){
        return new Promise(async (resolve, reject) => {
            const msalConfig = {
                auth: {
                    clientId: "",
                    authority: "https://login.microsoftonline.com/{TENANTID}",
                    clientSecret: "",
                    knownAuthorities: [],
                }
            }   
             
            const tokenRequest = {  
                scopes: ["https://outlook.office365.com/.default"]
            };  

            async function getToken(tokenRequest) {  
                const cca = new msal.ConfidentialClientApplication(msalConfig);  
                const msalTokenCache = cca.getTokenCache();  
                return await cca.acquireTokenByClientCredential(tokenRequest);  
            }
            
            async function authenticate() {  
                try {  
                    const authResponse = await getToken(tokenRequest);  
                    return authResponse.accessToken;  
                } catch (error) {  
                    console.log(error); 
                    reject(error); 
                }             
            };  
              
            let token = await authenticate();  
            resolve(token);
        });
    };

    EmailListener.init = async function () {
        const mailId = "youremail@website.com";  
        let token = await EmailListener.getAuthToken();
        let base64Encoded =  Buffer.from([`user=${mailId}`, `auth=Bearer ${token}`, '', ''].join('\x01'), 'utf-8').toString('base64');  

        let mailListener = new MailListener({
                xoauth2: base64Encoded,  
                host: 'outlook.office365.com',  
                port: 993,  
                tls: true,  
                debug: console.log,  
                authTimeout: 25000,  
                connTimeout: 30000,  
                tlsOptions: {  
                  rejectUnauthorized: false,  
                  servername: 'outlook.office365.com'  
                },
                mailbox: "INBOX", // mailbox to monitor
                searchFilter: ["UNSEEN"], // the search filter being used after an IDLE notification has been retrieved
                fetchUnreadOnStart: true, // use it only if you want to get all unread email on lib start. Default is `false`,
        });

        mailListener.start(); // start listening

        // stop listening
        //mailListener.stop();

        mailListener.on("server:connected", function () {
            console.log(Modules.System,"Email server connected");
        });

        mailListener.on("server:disconnected", function () {
            console.log("Email server disconnected, restarting...");
            EmailListener.init();
        });

        mailListener.on("error", function (err) {
            console.log(err);
        });

        mailListener.on("mail", async (mail, seqno, attributes) => {
            let html = mail.html;
            let body = mail.text;
            let from = mail.headers.from;
        });

        mailListener.on("attachment", function (attachment, email) {
            console.log(attachment.path);
        });

    };
}
