import { registerMgtComponents, Providers, SimpleProvider, ProviderState } from './mgt.js';
						
window.addEventListener("load", function() {
	registerMgtComponents();
	
	const picker = document.querySelector("#people-picker");
	const makeCall = document.querySelector("#make-call");
	const dialPads = document.querySelectorAll(".dial-key-wrap");
	
	for (let dialpad of dialPads) dialpad.addEventListener("click", (evt) => 
	{
		console.debug("dialpad clicked", evt.target.parentNode.getAttribute("data-key"));
		const destination = picker.shadowRoot.querySelector("fluent-text-field").value;
		picker.shadowRoot.querySelector("fluent-text-field").value = destination + evt.target.parentNode.getAttribute("data-key");
	});
	
	picker.addEventListener("selectionChanged", (evt) => {
		console.log("selected", evt.detail);

		if (evt.detail.length == 1) {
			dialSelected(evt.detail);
		}		
	})	
	
	makeCall.addEventListener("click", (evt) => 
	{		
		if (picker?.selectedPeople.length > 0) {
			dialSelected(picker.selectedPeople);			

		} else {
			dialPhoneNumber(picker);
		}
	});		
	
	Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);	
	Providers.globalProvider.login();
	
	picker.renderInput();
	picker.focus();	
});

window.addEventListener("unload", function() {
	
});

chrome.runtime.onMessage.addListener(async (msg) => {	
	console.debug("chrome.runtime.onMessage", msg);	
})

function dialPhoneNumber(picker) {
	const destination = picker.shadowRoot.querySelector("fluent-text-field").value;

	if (destination.length > 0) {	
		console.debug("input", destination); 
		chrome.runtime.sendMessage({action: "make_call", destination});	
		window.close();		
	} else { 
		alert("Unable to make call, target has not been selected or input provided");
	}	
}

function dialSelected(selected) {
	console.log("dialSelected", selected);

	if (selected.length == 1) {
		const person = selected[0];
		const displayName = person.displayName;			
		const destination = getDestination(person);

		if (destination) {
			chrome.runtime.sendMessage({action: "make_call", destination, displayName});	
			window.close();
		} else {
			alert("Unable to make call, target is not an organization user or has no telephone number");
		}				
	} 
	else
		
	if (selected.length > 1) { // TODO call list
		const contacts = [];
		
		for (let person of selected) {
			contacts.push({name: person.displayName, destination: getDestination(person), email: person.scoredEmailAddresses[0]?.address});
		}
		
		const url = "/call-list.html";
		const action = "open_webapp";
		const key = "call-list";	
		chrome.runtime.sendMessage({action,  data: {key, url, contacts}, width: 800, height: 800});	
		window.close();		
	}
}

function getDestination(person) {
	console.debug("getDestination", person);
	
	let destination = null;

	if (person.personType?.subclass == "OrganizationUser") {
		destination = person.id;
		
	}
	else 
	
	if (person.phones.length > 0) {
		destination = person.phones[0].number.replaceAll(" ", ""); // TODO pick list
	}
	else
		
	if (person.imAddress) {
		destination = person.imAddress;
		if (destination.indexOf("@") > -1) destination = destination.split("@")[0];
	}
	return destination;
}

async function getAccessToken(scopes) {
	console.debug("getAccessToken scope", scopes);	
	const url = getSetting("cas_server_url") + "/teams/api/openlink/msal/token/graph/" + getScopes(scopes);				
	const resp = await fetch(url, {method: "GET", headers: {authorization: getSetting("cas_server_token")}});	
	const json = await resp.json();
	console.debug("getAccessToken token", json);	
	return Promise.resolve(json.access_token);
}

function login() {
	console.debug("login");	
	Providers.globalProvider.setState(ProviderState.SignedIn)
}

function logout() {
	console.debug("logout");		
	Providers.globalProvider.setState(ProviderState.SignedOut)
}

function getScopes(scopes) {
	return "User.Read, User.ReadWrite, User.Read.All, People.Read, User.ReadBasic.All, presence.read.all, Mail.ReadBasic, Tasks.Read, Group.Read.All, Tasks.ReadWrite, Group.ReadWrite.All";
}

