import { registerMgtComponents, Providers, SimpleProvider, prepScopes, ProviderState } from './mgt.js';

let graphClient, activeTaskList, callList, picker, activeCall, nextCall;
						
window.addEventListener("load", function() {
	const casData = urlParam("cas_data");
	
	if (casData) {
		let data = JSON.parse(casData);
		console.debug("load", data);
		displayContacts(data.contacts);
	}
	
	registerMgtComponents();		
	Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);	
	Providers.globalProvider.login();	

	callList = document.querySelector("#call-list");	
	picker = document.querySelector("#people-picker");	
	
	picker.addEventListener("selectionChanged", async (evt) => {
		console.log("selected", evt.detail);
		
		if (evt.detail.length > 0) {
			const person = evt.detail[0];
			const contact = {name: person.displayName, destination: getDestination(person), email: person.scoredEmailAddresses[0]?.address};
		
			if (contact.destination) {
				const taskData = {
					title: contact.name, 
					body: {content: contact.email, contentType: "text"}, 
					linkedResources: [{webUrl: "tel:" + contact.destination, applicationName: "CAS Call Assistant", displayName: "comitFS"}]
				};
				
				const newTask = await graphClient.api("/me/todo/lists/" + activeTaskList + "/tasks").header('Cache-Control', 'no-store').post(taskData);
				
				if (!newTask) {
					alert("Call List cannot be updated");
				} else {
					addContact(contact, newTask);
				}
			} else {
				alert(contact.name + " does not have a telephone number");
			}
			
			picker.selectedPeople = [];			
		}
	})			
});

window.addEventListener("unload", function() {
	
});
chrome.runtime.onMessage.addListener(async (msg) => {	
	console.debug("chrome.runtime.onMessage", msg);	
		
	switch (msg.action) {		
		case "make_next_call":
		
			if (nextCall) {
				makeCall(nextCall);		
			}
			break;

		case "notify_cas_dialer_connected":

			if (activeCall) {			
				removeClass("make-call");					
				removeClass("progress-call");				
				activeCall.classList.add("in-call");
				activeCall.innerHTML = `
					<svg width="24" height="18" viewBox="0 0 18 18" fill="#ffffff" xmlns="http://www.w3.org/2000/svg">
					  <path d="M3.98706 1.06589C4.89545 0.792081 5.86254 1.19479 6.31418 2.01224L6.38841 2.16075L7.04987 3.63213C7.46246 4.54992 7.28209 5.61908 6.60754 6.3496L6.47529 6.48248L5.43194 7.45541C5.24417 7.63298 5.38512 8.32181 6.06527 9.49986C6.67716 10.5597 7.17487 11.0552 7.41986 11.0823L7.4628 11.082L7.5158 11.0716L9.56651 10.4446C10.1332 10.2713 10.7438 10.4487 11.1298 10.8865L11.2215 11.0014L12.5781 12.8815C13.1299 13.6462 13.0689 14.6842 12.4533 15.378L12.3314 15.5039L11.7886 16.018C10.7948 16.9592 9.34348 17.2346 8.07389 16.7231C6.13867 15.9433 4.38077 14.1607 2.78368 11.3945C1.18323 8.62242 0.519004 6.20438 0.815977 4.13565C0.99977 2.85539 1.87301 1.78674 3.07748 1.3462L3.27036 1.28192L3.98706 1.06589Z"></path>			
					</svg>In a Call&nbsp;
				`;	
				activeCall.disabled = true;					
			}
			break;
		
		case "notify_cas_dialer_disconnected":

			if (activeCall) {
				removeClass("make-call");
				removeClass("progress-call");					
				removeClass("in-call");
				activeCall.classList.add("completed-call");					
				activeCall.innerHTML = `
					<svg width="24" height="18" viewBox="0 0 18 18" fill="#ffffff" xmlns="http://www.w3.org/2000/svg">
					  <path d="M10 2a8 8 0 110 16 8 8 0 010-16zm3.36 5.65a.5.5 0 00-.64-.06l-.07.06L9 11.3 7.35 9.65l-.07-.06a.5.5 0 00-.7.7l.07.07 2 2 .07.06c.17.11.4.11.56 0l.07-.06 4-4 .07-.08a.5.5 0 00-.06-.63z" fill="currentColor"></path>
					</svg>&nbsp;Completed&nbsp;			
				`				
			}		
			break;
	}
})

function removeClass(className) {
	if (activeCall.classList.contains(className)) activeCall.classList.remove(className);
}

function urlParam(name) {
	var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	if (!results) { return undefined; }
	return unescape(results[1] || undefined);
};


async function getGraphClient() {
	graphClient = Providers.globalProvider.graph.client;
	let taskLists = await graphClient.api("/me/todo/lists").header('Cache-Control', 'no-store').get();
	console.debug("Graph API /me/todo/lists", taskLists);	
	
	const sortableLists = document.querySelector(".sortable-lists");
	
	for (let taskList of taskLists.value) {
		console.debug("getGraphClient", taskList);
		
		if (taskList.displayName.startsWith("CL:")) {
			const div = document.createElement("div");

			div.innerHTML = `
				<div class="">
				  <div class="">
					<div role="none" draggable="true">
					  <li class="listItem-container active" role="option" aria-selected="true" tabindex="-1" data-is-focusable="true">
						<div class="listItem">
						  <div class="listItem-inner">
							<span class="listItem-icon">
							  <svg class="fluentIcon ___12fm75w f1w7gpdv fez10in fg4l7m0" fill="currentColor" aria-hidden="true" width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg" focusable="false">
								<path d="M3 6.5a1 1 0 100-2 1 1 0 000 2zm3-1c0-.28.22-.5.5-.5h11a.5.5 0 010 1h-11a.5.5 0 01-.5-.5zm0 5c0-.28.22-.5.5-.5h11a.5.5 0 010 1h-11a.5.5 0 01-.5-.5zm.5 4.5a.5.5 0 000 1h11a.5.5 0 000-1h-11zm-2.5.5a1 1 0 11-2 0 1 1 0 012 0zm-1-4a1 1 0 100-2 1 1 0 000 2z" fill="currentColor"></path>
							  </svg>
							</span>
							<span class="listItem-title listItem-titleParsed">
							  <span id="${taskList.id}">${taskList.displayName.substring(3)}</span>
							</span>
							<span class="listItem-count" aria-hidden="true">1</span>
						  </div>
						</div>
					  </li>
					</div>
				  </div>
				</div>
			`;
			
			sortableLists.appendChild(div);
			
			div.addEventListener("click", async function(evt) {
				console.debug("div clicked", evt.target.id);				
				activeTaskList = evt.target.id
				document.getElementById("task-name").innerHTML = evt.target.innerHTML;
				
				if (activeTaskList) {
					const tasks = await graphClient.api("/me/todo/lists/" + activeTaskList + "/tasks").header('Cache-Control', 'no-store').get();					
					
					if (tasks) 	{
						callList.innerHTML = "";						
						
						for (let task of tasks.value) {
							const contact = {status: task.status, name: task.title, destination: task.linkedResources[0]?.webUrl.substring(4), email: task.body?.content};
							addContact(contact, task);
						}
			
						const taskDisplay = document.querySelector(".flex-container");
						taskDisplay.style.display = "";
					}
				}
			});		
		}
	}
	
	const addList = document.querySelector(".sidebar-addList");
	
	addList.addEventListener("click", async function() {
		let displayName = prompt("Enter Group Name");
		
		if (displayName) {
			displayName = "CL:" + displayName;
			const listData = { displayName };
			
			await graphClient.api('/me/todo/lists').header('Cache-Control', 'no-store').post(listData);
			location.reload();
		}
	});
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
	getGraphClient();	
}

function logout() {
	console.debug("logout");		
	Providers.globalProvider.setState(ProviderState.SignedOut)
}

function getScopes(scopes) {
	return "User.Read, User.ReadWrite, User.Read.All, People.Read, User.ReadBasic.All, presence.read.all, Mail.ReadBasic, Tasks.Read, Group.Read.All, Tasks.ReadWrite, Group.ReadWrite.All";
}

async function addContact(contact, task) {
	console.debug("addContact", contact, task);

	const status = contact.status ? contact.status : "notStarted";
	const contactRow = document.createElement("fluent-data-grid-row");
	const person = document.createElement("fluent-data-grid-cell"); 
	const action = document.createElement("fluent-data-grid-cell"); 		
	
	action.setAttribute("grid-column", 2);
	action.style.width = "150px";	
	
	person.setAttribute("grid-column", 1);
	person.style.width = "400px";
	
	person.innerHTML = `<mgt-person person-query="${contact.email}" line2-property="jobTitle" style="--person-line1-font-size: 24px;margin-bottom:5px;" view="threelines" show-presence person-presence='{"availability": "Offline", "activity": "Offline"}' fallback-details='{"displayName":"${contact.name}"}'></mgt-person>`;			

	if (status == "completed") {
		await graphClient.api("/me/todo/lists/" + activeTaskList + "/tasks/" + task.id).header('Cache-Control', 'no-store').delete();
		return
	}
	else 
		
	if (status == "notStarted") {	
		action.innerHTML = `
				<fluent-button id="${task.id}" class="make-call" appearance="neutral" title="Make Call to ${contact.destination}" data-destination="${contact.destination}" data-name="${contact.name}">
					<svg width="24" height="16" viewBox="0 0 16 16" fill="#ffffff" xmlns="http://www.w3.org/2000/svg">
					  <path d="M3.98706 1.06589C4.89545 0.792081 5.86254 1.19479 6.31418 2.01224L6.38841 2.16075L7.04987 3.63213C7.46246 4.54992 7.28209 5.61908 6.60754 6.3496L6.47529 6.48248L5.43194 7.45541C5.24417 7.63298 5.38512 8.32181 6.06527 9.49986C6.67716 10.5597 7.17487 11.0552 7.41986 11.0823L7.4628 11.082L7.5158 11.0716L9.56651 10.4446C10.1332 10.2713 10.7438 10.4487 11.1298 10.8865L11.2215 11.0014L12.5781 12.8815C13.1299 13.6462 13.0689 14.6842 12.4533 15.378L12.3314 15.5039L11.7886 16.018C10.7948 16.9592 9.34348 17.2346 8.07389 16.7231C6.13867 15.9433 4.38077 14.1607 2.78368 11.3945C1.18323 8.62242 0.519004 6.20438 0.815977 4.13565C0.99977 2.85539 1.87301 1.78674 3.07748 1.3462L3.27036 1.28192L3.98706 1.06589Z"></path>			
					</svg>Call Contact&nbsp;
				</fluent-button>		
		`;	
	}
	
	action.addEventListener("click", async (evt) => {
		makeCall(evt.target);
	});		
	
	contactRow.appendChild(person);
	contactRow.appendChild(action);		
	callList.appendChild(contactRow);
	return null;
}

async function makeCall(node) {
	const destination = node.getAttribute("data-destination");			
	const displayName = node.getAttribute("data-name");	
	
	nextCall = node.parentNode.parentNode.nextElementSibling?.querySelector("fluent-button");
	console.debug("button clicked", destination, displayName, node.id, nextCall);

	chrome.runtime.sendMessage({action: "make_call", destination, displayName});			
	
	const updateData = {status: "completed"};
	const response = await graphClient.api("/me/todo/lists/" + activeTaskList + "/tasks/" + node.id).header('Cache-Control', 'no-store').patch(updateData);
	activeCall = document.getElementById(node.id);
	removeClass("make-call");			
	activeCall.classList.add("progress-call");	
	activeCall.innerHTML = `
		<svg width="24" height="18" viewBox="0 0 18 18" fill="#ffffff" xmlns="http://www.w3.org/2000/svg">
		  <path d="M3.98706 1.06589C4.89545 0.792081 5.86254 1.19479 6.31418 2.01224L6.38841 2.16075L7.04987 3.63213C7.46246 4.54992 7.28209 5.61908 6.60754 6.3496L6.47529 6.48248L5.43194 7.45541C5.24417 7.63298 5.38512 8.32181 6.06527 9.49986C6.67716 10.5597 7.17487 11.0552 7.41986 11.0823L7.4628 11.082L7.5158 11.0716L9.56651 10.4446C10.1332 10.2713 10.7438 10.4487 11.1298 10.8865L11.2215 11.0014L12.5781 12.8815C13.1299 13.6462 13.0689 14.6842 12.4533 15.378L12.3314 15.5039L11.7886 16.018C10.7948 16.9592 9.34348 17.2346 8.07389 16.7231C6.13867 15.9433 4.38077 14.1607 2.78368 11.3945C1.18323 8.62242 0.519004 6.20438 0.815977 4.13565C0.99977 2.85539 1.87301 1.78674 3.07748 1.3462L3.27036 1.28192L3.98706 1.06589Z"></path>			
		</svg>Dialing....&nbsp;
	`;	
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