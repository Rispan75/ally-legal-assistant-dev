// Description: This file contains the code to run the task pane add-in.
// Dependencies: This file depends on the following files:
//   1. src/taskpane/taskpane.css

import { ask } from './Containers/aks.js'; 
import { summary } from './Containers/summary_selected.js';
import { document_summary } from './Containers/summary_document.js';

import { createNestablePublicClientApplication } from "@azure/msal-browser";
// import { PublicClientApplication } from "@azure/msal-browser"      //Use this line instead of above line, if getting error: {"error":{"code":"UserError","message":"Failed to render jinja template. Please modify your prompt to fix the issue."}}

let pca = undefined;

fetch("assets/config.json")
  .then((res) => res.text())
  .then((text) => {
    console.log("Config: ", text);
    const config = JSON.parse(text);
    localStorage.setItem('pfendpoint', config['prompt-flow-endpoint']);
    localStorage.setItem('pfconfigendpoint', config['prompt-flow-config-endpoint']);
    localStorage.setItem('clientId', config['clientId']);
    localStorage.setItem('authority', config['authority']);
    localStorage.setItem('sso-enabled', config['sso-enabled']);
     
   })
  .catch((e) => console.error(e));

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "block"; 
    const filename = Office.context.document.url.split('\\').pop().split('/').pop() //fetching filename from Word API
    localStorage.setItem('filename', filename);

    
    
    
  if (localStorage.getItem('sso-enabled') == "true"){
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: localStorage.getItem('clientId'),
        authority: localStorage.getItem('authority'),
      },
    });
  }

    getOpenAIResponseDemo(localStorage.getItem('pfendpoint')).then((result) =>
    {
      // write the name of the user based on the profile from SSO
      const name = localStorage.getItem("profile") ? JSON.parse(localStorage.getItem("profile")).displayName : "User";
      
      const welcomeMessage = document.getElementById("title-with-name");
      welcomeMessage.textContent = `Hello ${name}, ${welcomeMessage.textContent}`;
      console.log("Result: ", result);
      if (result != null) {
        setTimeout(() => {          
          document.getElementById("sideload-msg").style.display = "none";
          document.getElementById("app-body").style.display = "flex";
          document.getElementById("ask-button").onclick = ask;
          document.getElementById("index-doc-button").onclick = index_document;
          document.getElementById("fetchPolicyData").onclick = summary;
          document.getElementById("fetchSummaryData").onclick = document_summary;
          document.getElementById("reset-button").onclick = reset_cache;
          document.getElementById("iteration-button").onclick = iteration_logic;
        }, 1000);
      }
      else
      {
        showErrorMessage("An unexpected error occurred: " + result);
      }
    }).catch((error) => {
      console.error("Error messgae: " + error);
      showErrorMessage(error);
    })
    
  }
});

async function sso() {
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["User.Read", "openid", "profile"],
  };
  let accessToken = null;

  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }

  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }

  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    console.error(`Unable to acquire access token.`);
    return;
  }

  // Call the Microsoft Graph API with the access token.
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/memberOf?$select=displayName,id,description,mail,mailNickName,userPrincipalName`,
    {
      headers: { Authorization: accessToken },
    }
  );

  const response_profile = await fetch(
    `https://graph.microsoft.com/v1.0/me`,
    {
      headers: { Authorization: accessToken },
    }
  );


  if (response.ok && response_profile.ok) {
    // Write file names to the console.
    const me = await response_profile.json();
    // save to global variable for later use
    localStorage.setItem('profile', JSON.stringify(me));
    console.log("Profile: ", me);

    const data = await response.json();
    const groups = data.value.map((item) => item.id);
    localStorage.setItem('groups', JSON.stringify(groups));
    console.log("Groups: ", groups);
  }

  }

function showErrorMessage(message) {
  const sideloadMsg = document.getElementById("sideload-msg");

  // Update the content of the sideload message
  sideloadMsg.innerHTML = `<h1>There  been a connection check error</h1><p>The following error occurred: </p><p>${message}</p>`;

  // Style the error message
  sideloadMsg.style.display = "block";
  sideloadMsg.style.backgroundColor = "#ffffff";
  sideloadMsg.style.padding = "10px";
  sideloadMsg.style.border = "1px solid #f5c6cb";
  sideloadMsg.style.borderRadius = "5px";

  // hide after 5 seconds
  setTimeout(() => {
    sideloadMsg.style.display = "none";
  }, 5000);

}


export async function reset_cache() {
  localStorage.removeItem('FullSummaryData');
  localStorage.removeItem('groups');
  localStorage.removeItem('profile');
  showSuccessMessage("Cache has been reset successfully");
}


// Index document function - Demo now but will be implemented in the future with Azure Search, Prompt Flow
export async function index_document() {
  const spinner = document.getElementById("index-doc-spinner");
  const reviewContainer = document.getElementById("index-doc-container");
  const indexbutton = document.getElementById("index-doc-button"); 
  // Show spinner and hide container
  spinner.style.display = "flex";
  indexbutton.style.display = "none"; 
  reviewContainer.style.display = "block"; 

  // ✅ Change background color to light green

  reviewContainer.style.backgroundColor = "#e6f5e6"; // light green
  reviewContainer.style.border = "1px solid #ffffffff"; // green border
  reviewContainer.style.borderRadius = "10px";



  // Update heading to show "Indexing in progress..."

  const heading = reviewContainer.querySelector("h2");
  if (heading) {
    heading.textContent = "⏳Preparing Document";
    heading.style.color = "black";
    heading.style.fontWeight = "bold";

  }



  // Hide all <p> tags temporarily

  const pTags = reviewContainer.querySelectorAll("p");
  pTags.forEach((p) => {p.style.display = "none"; p.style.color = "black"; });



  try {

    // Get function-app endpoint

    const contract_index_endpoint = localStorage.getItem('contractindexendpoint');
    const filename = localStorage.getItem('filename');
    const response = await fetch(contract_index_endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },

      body: JSON.stringify({
        filename: filename
      })
    });

    const message = await response.text(); 

    if (response.ok) {
      console.log("Response: ok");
      // 1. Change the heading with a professional icon (balloon or checkmark)
      const heading = reviewContainer.querySelector('h2');
      if (heading) {
        heading.innerHTML = '✅Document Ready'; 
      }

      // 2. Light success styling for the whole container
      reviewContainer.style.backgroundColor = '#d4edda';     // Light green (success)
      reviewContainer.style.border = '1px solid #c3e6cb';     // Success border
      reviewContainer.style.borderRadius = '8px';
      reviewContainer.style.padding = '16px';
      reviewContainer.style.color = '#155724';                // Dark green text for success

      // 3. Update the first paragraph

      const pTags = reviewContainer.querySelectorAll('p');
      if (pTags.length > 0) {
        pTags[0].innerHTML = 'Your document is prepared. <br> 🎉 Your tool is now ready-to-use.';
        pTags[0].style.display = 'block';
        pTags[0].style.fontWeight = '500';

      }



      // 4. Remove the second paragraph if present

      if (pTags.length > 1) {
        pTags[1].remove();
      }



      // 5. Remove spinner and button

      if (spinner) spinner.remove();
      if (indexbutton) indexbutton.remove();



    } else {
      // 1. Change the heading
      console.log("Response: bad");
      const heading = reviewContainer.querySelector('h2');
      if (heading) {
        heading.textContent = '❌Document Not Ready';
      }



      // 1. Light red background for the entire review container

      reviewContainer.style.backgroundColor = '#f8d7da';  // Light red
      reviewContainer.style.border = '1px solid #f5c6cb'; // Border similar to Bootstrap danger alert
      reviewContainer.style.borderRadius = '8px';
      reviewContainer.style.padding = '16px';

      // 2. Modify the first paragraph with an outlined inner box for the message
      const pTags = reviewContainer.querySelectorAll('p');
      if (pTags.length > 0) {
        pTags[0].innerHTML = `
          Your document could not be prepared. It could be due to the following issue:<br>
          <div style="border: 1px solid #f5c6cb; background-color: #fef2f2; padding: 10px; border-radius: 4px; margin-top: 8px;">
            ${message}
          </div>`;
        pTags[0].style.display = 'block';
        pTags[0].style.color = '#721c24';  // Text color for error
      }



      // 3. Remove the second paragraph if it exists

      if (pTags.length > 1) {

        pTags[1].remove();

      }



      // 4. Spinner and button visibility

      spinner.style.display = "none";

      indexbutton.style.display = "block";

    }

  } 

  catch (error) {

    console.log("%cNetwork or server error: " + error.message, "color: red");

    // show to user

    const heading = reviewContainer.querySelector('h2');

    if (heading) {

      heading.textContent = '❌Document Not Ready';

    }



    // 1. Light red background for the entire review container

    reviewContainer.style.backgroundColor = '#f8d7da';  // Light red

    reviewContainer.style.border = '1px solid #f5c6cb'; // Border similar to Bootstrap danger alert

    reviewContainer.style.borderRadius = '8px';

    reviewContainer.style.padding = '16px';



    // 2. Modify the first paragraph with an outlined inner box for the message

    const pTags = reviewContainer.querySelectorAll('p');

    if (pTags.length > 0) {

      pTags[0].innerHTML = `

        Your document could not be prepared. It could be due to the following issue:<br>

        <div style="border: 1px solid #f5c6cb; background-color: #fef2f2; padding: 10px; border-radius: 4px; margin-top: 8px;">

          Error connecting to Function app.

        </div>`;

      pTags[0].style.display = 'block';

      pTags[0].style.color = '#721c24';  // Text color for error

    }



    // 3. Remove the second paragraph if it exists

    if (pTags.length > 1) {

      pTags[1].remove();

    }



    // 4. Spinner and button visibility
    spinner.style.display = "none";
    indexbutton.style.display = "block";
  } 
  finally {
    console.log("End of Indexing button function.");
  }
}

// Function to display a success message on the top ribbon
function showSuccessMessage(message) {
  const ribbon = document.querySelector('.warning-ribbon');
  const ribbonText = document.getElementById('ribbon-text');

  ribbonText.textContent = message;

  ribbon.style.display = "block";
  ribbon.classList.add("fade-in");
  ribbon.classList.remove("fade-out");
  setTimeout(() => {
    ribbon.classList.remove("fade-in");
    ribbon.classList.add("fade-out");
  }, 2000); 
  setTimeout(() => {
    ribbon.style.display = "none";
  }, 3000); 
  
}

async function getOpenAIResponseDemo(pfuri)
{
  // run sso function to get the profile and groups is SSO is enabled else use the demo profile and groups
  if ((localStorage.getItem('profile') == null && localStorage.getItem('groups') == null) || localStorage.getItem('sso-enabled') == "true")
    {
      try {
        if (localStorage.getItem('sso-enabled') == "true"){
          console.log("SSO enabled");
          await sso();
        }
        else
        {
          console.log("SSO not enabled");
          const profile = JSON.stringify(
          {
            "businessPhones": [],
            "displayName": "John Doe",
            "givenName": "John",
            "jobTitle": "Legal professional",
            "mail": "jhon.doe@microsoft.com",
            "mobilePhone": null,
            "officeLocation": "",
            "preferredLanguage": null,
            "surname": "Doe",
            "userPrincipalName": "john.doe@microsoft.com",
            "id": "877e9802-b713-4250-8701-c70d2c1e9a42"
        })
          console.log("Profile: ", profile);
          localStorage.setItem('profile', profile);
          
          //localStorage.setItem('groups',['2846190d-05dc-4048-90bc-7e236f34d84b','62edbd7b-8d46-4d2c-a5a1-da5b78ba1d38','be8ca378-bc74-46c1-b922-e7f552486ede']);
          localStorage.setItem('groups', JSON.stringify(['22a229bd-c7a2-49d0-9eaa-e1fc888daac6','2846190d-05dc-4048-90bc-7e236f34d84b','62edbd7b-8d46-4d2c-a5a1-da5b78ba1d38','be8ca378-bc74-46c1-b922-e7f552486ede']));

          //localStorage.setItem('groups', JSON.stringify(['22a229bd-c7a2-49d0-9eaa-e1fc888daac6']));

          console.log("Profile: ", JSON.parse(localStorage.getItem('profile')));
          console.log("Groups: ", localStorage.getItem('groups'));
        }
      }
      catch (error) {
        return error;
      }
    }
    else
    {
      console.log("Profile already exists");
      console.log("Profile: ", JSON.parse(localStorage.getItem('profile')));
      console.log("Groups: ", localStorage.getItem('groups'));
      
      
    }
  
  const uri = new URL(pfuri).origin
  
  checkdocumentindex()
  
  return "Success";
}

// action on change of language-select
document.getElementById("language-select").onchange = function() {
  var lang = document.getElementById("language-select").value;
  localStorage.setItem('language', lang);
  console.log("Language: ", lang);
}

async function checkdocumentindex()
{
  // check if the document has been indexed
  console.log("check index")
  console.log(localStorage.getItem('filename'));
  const response = await fetchData(localStorage.getItem('pfendpoint'), localStorage.getItem('filename'), localStorage.getItem('groups'));       
  const data = await response.json(); 
  console.log(data.answer.Found);
  
  if (data.answer.Found == false)
  {
    document.getElementById("index-doc-container").style.display = "flex";    
    //change lebel filename-notindexed-label to the filename
    document.getElementById("filename-notindexed-label").textContent = localStorage.getItem('filename');
    
  }
}


async function fetchData(endpoint, filename, groups ) {  
  return await fetch(endpoint, {  
      method: 'POST',  
      headers: {  
          'Content-Type': 'application/json'  
      },  
      body: JSON.stringify({  
          query_type: 99,
          filename: filename,
          groups: JSON.parse(groups)        
      })  
  });  
}  


export async function iteration_logic() {  

  const spinner = document.getElementById("iteration-spinner");

  const reviewContainer = document.getElementById("iteration-container");

  const indexbutton = document.getElementById("iteration-button");

 

 

  // Show spinner and hide container

  spinner.style.display = "flex";

  indexbutton.disabled = true;

  indexbutton.classList.add("disabled-style");

 

  reviewContainer.style.display = "block";

 

 

  // Get function-app endpoint

  const pfendpoint = localStorage.getItem('pfendpoint');

  const filename = localStorage.getItem('filename');

  const language = localStorage.getItem('language');

 

 

  const response = await fetch(pfendpoint, {

    method: 'POST',

    headers: {

      'Content-Type': 'application/json'

    },

    body: JSON.stringify({

      query_type: 10,

      filename: filename,

      language: language,

    })

  });

 

  const message = await response.text();

  console.log(message);

 

  const pmessage = JSON.parse(message); // ✅ Use a different variable name

 

  indexbutton.disabled = false;
  indexbutton.classList.remove("disabled-style");

  if (response.ok) {
    reviewContainer.style.display = "block";
    spinner.style.display = "none";

let formatted = `<h2 style="margin-bottom: 20px; font-size: 22px; font-weight: 600; color: #2c3e50;">Unused Policies</h2>`;
pmessage.answer.forEach((item, index) => {
  formatted += `
    <div style="
      margin-bottom: 24px;
      border-radius: 16px;
      background-color: #ffffff;
      box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
      overflow: hidden;
    ">
      <div style="
        background: linear-gradient(90deg, #3a8bfd, #2c52ff);
        color: #ffffff;
        font-weight: bold;
        padding: 14px 18px;
        font-size: 16px;
      ">
        Policy - ${index + 1}

      </div>

      <div style="
        padding: 20px;
        display: flex;
        flex-direction: column;
        gap: 16px;
      ">


        <div style="

          border: 1px solid #e0e0e0;

          background-color: #fafafa;

          padding: 12px;

          border-radius: 10px;

        ">

          <div style="font-weight: bold; color: #004aad; margin-bottom: 4px;">Title</div>

          <div style="color: #333;">${item.Title}</div>

        </div>

 

        <div style="

          border: 1px solid #e0e0e0;

          background-color: #fafafa;

          padding: 12px;

          border-radius: 10px;

        ">

          <div style="font-weight: bold; color: #004aad; margin-bottom: 4px;">Summary</div>

          <div style="color: #333;">${item.summary}</div>

        </div>

      </div>

    </div>

  `;

});

 

 

 

    document.getElementById("iteration-container").innerHTML = formatted;

  } else {

    reviewContainer.style.display = "block";

    spinner.style.display = "none";

    document.getElementById("iteration-container").innerText = "Operation Failed";

 

  }

}
