import fetch from 'node-fetch';
import { Headers } from 'node-fetch';
import dotenv from 'dotenv';
dotenv.config();
import readline from 'readline';

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;

const userToNotShow = [
    "AdeleV@yxr7j.onmicrosoft.com",
    "AlexW@yxr7j.onmicrosoft.com",
    "DiegoS@yxr7j.onmicrosoft.com",
    "GradyA@yxr7j.onmicrosoft.com",
    "HenriettaM@yxr7j.onmicrosoft.com",
    "IsaiahL@yxr7j.onmicrosoft.com",
    "JohannaL@yxr7j.onmicrosoft.com",
    "JoniS@yxr7j.onmicrosoft.com",
    "LeeG@yxr7j.onmicrosoft.com",
    "LidiaH@yxr7j.onmicrosoft.com",
    "LynneR@yxr7j.onmicrosoft.com",
    "MeganB@yxr7j.onmicrosoft.com",
    "MiriamG@yxr7j.onmicrosoft.com",
    "NestorW@yxr7j.onmicrosoft.com",
    "PattiF@yxr7j.onmicrosoft.com",
    "PradeepG@yxr7j.onmicrosoft.com",

]
const userToShow = [
    "AdeleV@yxr7j.onmicrosoft.com",
]

async function patchAdUser(headers, id, body) {
    // const response = await fetch(`https://graph.microsoft.com/v1.0/users/${id}`, {headers: headers, method: 'PATCH', body: JSON.stringify(body)});
    // //log if error
    // if (!response.ok) {
    //     console.log(response);
    // }
    const response = await fetch(`https://graph.microsoft.com/v1.0/users/${id}`,
    {
        method: 'PATCH',
        headers: headers,
        body: JSON.stringify(body)
    });
    // const data = await response.json();
    if (response.ok) {
        console.log(`L'utilisateur avec l'ID ${id} a été mis à jour.`);
    } else {
        console.error(`La mise à jour de l'utilisateur avec l'ID ${id} a échoué. Statut de la réponse : ${response.status}`);
    }
    return response;
}

async function getAccesToken(){
    let details = {
        "grant_type": 'client_credentials',
        "client_id": clientId,
        "client_secret": clientSecret,
        "scope": 'https://graph.microsoft.com/.default',
    };
    
    let formBody = [];
    for (let property in details) {
      let encodedKey = encodeURIComponent(property);
      let encodedValue = encodeURIComponent(details[property]);
      formBody.push(encodedKey + "=" + encodedValue);
    }
    formBody = formBody.join("&");
    const met = {
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    const headerss = new Headers(met);
    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: headerss,
        body: formBody
    })
    const token = await response.json();
    return token.access_token;
}


async function main(){
    const token = await getAccesToken();
    const meta = {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token
    };
    const body = {
        [`extension_7ab43b13719b4cbbbfb83787a26e7952_${process.env.EXTENSION_NAME}`]: true
    }
    console.log(JSON.stringify(body));
    const headers = new Headers(meta);

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
     });
    rl.question("Tapez 1 pour afficher dans le trombinoscope les utilisateurs indiquer dans la variable userToShow, tapez 2 pour ne pas afficher les utilisateurs dans userToNotShow\n", function(input) {
        switch (input) {
            case "1":
                userToShow.forEach(async (user) => {
                    await patchAdUser(headers, user, body);
                });
                break;
            case "2":
                userToNotShow.forEach(async (user) => {
                    await patchAdUser(headers, user, body);
                });
                break;
            default:
                console.log("Vous n'avez pas choisi une option valide");
                break;
        }
        rl.close();
    });

    // userToShow.forEach(async (user) => {
    //     await patchAdUser(headers, user, body);
    // });
}

main();
