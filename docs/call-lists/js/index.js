import { TemplateHelper, registerMgtComponents, Providers, Msal2Provider, SimpleProvider, ProviderState } from './mgt.js';

let callId, callView, clientId, searchView, viewClient, selectedContact, telCache, currentCli, holdCall, transferCall, inviteUser, inviteToMeeting, acceptCall, declineCall, endCall, returnCall, requestToJoin, nextCall, acceptSuggestion, declineSuggestion, internalCollab, clearCache, assistButton, assistText, callOptions, callControls;
						
window.addEventListener("load", function() {
	const data = JSON.parse(urlParam("data"));
	console.debug("load", data);

	document.title = data.type == "incoming" ? "Incoming Call" : "Outgoing Call";
	callId = data.id;
	registerMgtComponents();
	
	currentCli = data.callerId;	
	telCache = localStorage.getItem("cas.telephone.cache");
	if (!telCache) telCache = "{}";
	telCache = JSON.parse(telCache);	
	
	const dialPad = document.querySelector("#dial-pad");
	dialPad.shadowRoot.querySelector("button").style.backgroundColor = "white";
	
	callView = document.querySelector(".call-view");	
	searchView = document.querySelector(".call-search");		
	clientId = document.querySelector("#client-id");
	
	viewClient = document.querySelector("#view-client");
	acceptCall = document.querySelector("#accept-call");
	declineCall	= document.querySelector("#decline-call");
	holdCall = document.querySelector("#hold-call");
	transferCall = document.querySelector("#transfer-call");
	inviteUser	= document.querySelector("#invite-user");	
	inviteToMeeting = document.querySelector("#invite-to-meeting");
	
	acceptSuggestion = document.querySelector("#accept-suggestion");
	declineSuggestion = document.querySelector("#decline-suggestion");
	
	endCall = document.querySelector(".end-call");
	internalCollab = document.querySelector(".internal-collab");
	clearCache = document.querySelector(".clear-cache");
	callOptions = document.querySelector(".call-options");
	
	assistButton = document.querySelector("#assist-button");
	assistText = document.querySelector("#assist-text");
	callControls = document.querySelector("#call-controls");
	
	viewClient.addEventListener("click", (evt) => {
		handleViewClientAction(evt.target);
	})
	
	nextCall = document.querySelector("#make-next-call");	

	nextCall.addEventListener("click", (evt) => {
		chrome.runtime.sendMessage({action: "make_next_call"});	
		//window.close();
	});		
	
	returnCall = document.querySelector("#return-call");	
	requestToJoin = document.querySelector("#request-join-call");
	
	returnCall.addEventListener("click", (evt) => {
		chrome.runtime.sendMessage({action: "make_call", destination: data.callerId});
		returnCall.style.display = "none";	
		endCall.style.display = "none";	
		viewClient.style.display = "";	
	})	

	requestToJoin.addEventListener("click", (evt) => {
		requestToJoin.style.display = "none";	
		const destination = getSetting("cas_delegate_userid", "ba9e081a-5748-40ca-8fd5-ab9c74dae3d1");
		const data = {request: {features: {callerName: selectedContact.displayName,	calledId: destination}}};
		requestToJoinCall(data);
	})	
	
	acceptSuggestion.style.display = "none";	
	declineSuggestion.style.display = "none";
		
	acceptSuggestion.addEventListener("click", () => {
		viewClient.innerHTML = "View Client";	
		viewClient.id = "view";	
		acceptSuggestion.style.display = "none";	
		declineSuggestion.style.display = "none";		
	});
	
	declineSuggestion.addEventListener("click", () => {
		viewClient.id = "search";
		doSearch();
		acceptSuggestion.style.display = "none";	
		declineSuggestion.style.display = "none";		
	});	

	clearCache.addEventListener("click", () => {
		localStorage.removeItem("cas.telephone.cache"); // hack to clear cache
	});	
	
	internalCollab.addEventListener("click", () => {
		const chatId = getSetting("cas_active_call_thread_id", "19:83ec482c-3bc5-4116-acee-e081cc720630_ba9e081a-5748-40ca-8fd5-ab9c74dae3d1@unq.gbl.spaces")
		chrome.runtime.sendMessage({action: "open_mgt_chat", url: getSetting("cas_server_url") + "/teams/mgt-chat/?chatId=" + chatId });
	});	
	
	assistButton.addEventListener("click", async (evt) => {
		const promptText = assistText.value;
		const url = getSetting("cas_server_url") + "/teams/api/openlink/prompt";				
		const resp = await fetch(url, {method: "POST", body: promptText, headers: {authorization: getSetting("cas_server_token")}});	
		const answer = await resp.text();
		assistText.value = "Question: " + promptText + "\n\n" + "Answer: " + answer;
	});
	
	callOptions.addEventListener("click", async (evt) => {
		viewClient.style.display = "none";	
		acceptCall.style.display = "none";
		declineCall.style.display = "none";	
		returnCall.style.display = "none";
		nextCall.style.display = "none";		
		requestToJoin.style.display = "none";			
		callControls.style.display = "";			
	});
	
	acceptCall.style.display = "none";
	declineCall.style.display = "none";			
	endCall.style.display = "none";	
	returnCall.style.display = "none";
	nextCall.style.display = "none";	
	requestToJoin.style.display = "none";		
	callControls.style.display = "none";		
	
	if (getSetting("cas_use_active_call_control", true)) {
		viewClient.style.display = "none";	

		if (data.type == "incoming") {
			acceptCall.style.display = "";
			declineCall.style.display = "";	
		}
		
		endCall.style.display = "";
		
		acceptCall.addEventListener("click", (evt) => {
			chrome.runtime.sendMessage({action: "accept_call", id: callId});
		})
		
		declineCall.addEventListener("click", (evt) => {
			chrome.runtime.sendMessage({action: "reject_call", id: callId});
		})
			
		endCall.addEventListener("click", (evt) => {
			chrome.runtime.sendMessage({action: "hangup_call", id: callId});
			setTimeout(() => chrome.runtime.sendMessage({action: "set_presence_dnd"}), 3000);	
		})
		
		holdCall.addEventListener("click", (evt) => 
		{
			if (clientId.classList.contains("call-held")) {
				chrome.runtime.sendMessage({action: "resume_call", id: callId});				
			} else {
				chrome.runtime.sendMessage({action: "hold_call", id: callId});
			}
		})
		
		transferCall.addEventListener("click", (evt) => {
			const picker = document.querySelector("#call-people-picker");
			
			if (picker?.selectedPeople.length > 0) {
				const person = picker.selectedPeople[0];
				let destination = getDestination(person);

				if (destination) {
					chrome.runtime.sendMessage({action: "transfer_call", id: callId, destination});				
				} else {
					alert("Unable to transfer, target is not an organization user or has no telephone number");
				}				
			} else {
				const destination = picker.shadowRoot.querySelector("fluent-text-field").value;
				
				if (destination.length > 0) {	
					console.debug("input", destination); 
					chrome.runtime.sendMessage({action: "transfer_call", id: callId, destination});					
				} else { 
					alert("Unable to transfer call, target has not been selected or input provided");
				}
			}
		})
			
		inviteUser.addEventListener("click", (evt) => {			
			const picker = document.querySelector("#call-people-picker");
			
			if (picker?.selectedPeople.length > 0) {
				const person = picker.selectedPeople[0];
				let destination = getDestination(person);

				if (destination) {
					chrome.runtime.sendMessage({action: "add_third_party", id: callId, destination});
				} else {
					alert("Unable to add participant to call, target is not an organization user or has no telephone number");
				}
				
			} else {
				const destination = picker.shadowRoot.querySelector("fluent-text-field").value;
				
				if (destination.length > 0) {	
					console.debug("input", destination); 
					chrome.runtime.sendMessage({action: "add_third_party", id: callId, destination});					
				} else { 
					alert("Unable to add participant to call, target has not been selected or input provided");
				}
			}
		})	
		
		inviteToMeeting.addEventListener("click", (evt) => {
			const key = clientId.personDetails.id;
			const emails = clientId.personDetails.scoredEmailAddresses;
			const name = clientId.personDetails.displayName;
			
			joinMeeting(key, emails, name, data);			
		})
		
	}		
	
	selectedContact = getClient(data);		
	clientId.fallbackDetails.displayName = selectedContact.displayName;
		
	if (data.emailAddress) {
		clientId.personQuery = data.emailAddress;
	} else if (selectedContact?.scoredEmailAddresses) {
		clientId.personQuery = selectedContact.scoredEmailAddresses[0].address;
	}		
	
	changeCallPresence("Available", "Available");	
	callView.style.display = "";	
	
	Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);	
	Providers.globalProvider.login();	
});

window.addEventListener("unload", function() {
	
});

chrome.runtime.onMessage.addListener(async (msg) => {	
	console.debug("chrome.runtime.onMessage", msg);	
	
	if (getSetting("cas_use_active_call_control", true)) {	
		//chrome.runtime.sendMessage({action: "set_presence_dnd"});
	}
			
	switch (msg.action) {
		case "update_active_call":
			if (msg.data.type == "elsewhere") {
				handleElseWhereCall(msg);					
				resizeWindowMinimised();
			}
			else
			if (msg.data.type == "missed") {
				handleMissedCall();				
				resizeWindowMinimised();				
			}				
			break;
			
		case "notify_shared_active_call":
			handleSharedActiveCall(msg);
			break;
			
		case "request_join_call":			// notification click
			requestToJoinCall(msg);
			break
			
		case "notify_cas_dialer_connected":
			handleConnectedCall(msg);		
			break;
			
		case "notify_cas_dialer_conferenced":
			handleConferencedCall(msg);		
			break;
			
		case "notify_cas_dialer_held":
			handleHeldCall(msg);		
			break;		
		case "notify_cas_dialer_disconnected":
			handleDisconnectedCall(msg);			
			break;
	}
})

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
	callView.style.display = "";
}

function logout() {
	console.debug("logout");		
	Providers.globalProvider.setState(ProviderState.SignedOut)
}

function getScopes(scopes) {
	return "User.Read, User.ReadWrite, User.Read.All, People.Read, User.ReadBasic.All, presence.read.all, Mail.ReadBasic, Tasks.Read, Group.Read.All, Tasks.ReadWrite, Group.ReadWrite.All";
}

function requestToJoinCall(data) {
	console.debug("requestToJoinCall", data);

	const message = "Can you please invite me to the call with " + data.request.features.callerName;
	const authorization = getSetting("cas_server_token");
	const url = getSetting("cas_server_url") + "/teams/api/openlink/chatThreadId/" + data.request.features.calledId + "/" + message;			
	fetch(url, {method: "POST", headers: {authorization}});	
}

function changeCallPresence(activity, availability) {
	clientId.showPresence = false;
	clientId.personPresence.activity = activity;
	clientId.personPresence.availability = availability;
	clientId.showPresence = true;	
}

function handleDisconnectedCall(data) {
	changeCallPresence("Offline", "Offline");
	clientId.classList.remove("call-connected");
	clientId.classList.add("call-disconnected");	

	if (document.title.startsWith("Outgoing")) {
		nextCall.style.display = "";	
	}
	
	returnCall.style.display = "";
	requestToJoin.style.display = "none";		
	acceptCall.style.display = "none";
	declineCall.style.display = "none";	
	endCall.style.display = "none";	
	callControls.style.display = "none";	
	hideClock();
}

function handleMissedCall(data) {
	changeCallPresence("Offline", "Offline");	
	
	returnCall.style.display = "";	
	requestToJoin.style.display = "none";		
	acceptCall.style.display = "none";
	declineCall.style.display = "none";		
	endCall.style.display = "none";
	
	if (viewClient.id == "view") {
		//viewClient.style.display = "";		
	}	
}

function handleElseWhereCall(data) {
	changeCallPresence("Away", "Away");	

	requestToJoin.style.display = "";		
	returnCall.style.display = "none";		
	acceptCall.style.display = "none";
	declineCall.style.display = "none";	
	endCall.style.display = "none";
	viewClient.style.display = "";		
}

function handleHeldCall(data) {
	clientId.classList.remove("call-connected");	
	clientId.classList.add("call-held");	
}

function handleConnectedCall(data) {
	callId = data.id;
	clientId.classList.remove("call-disconnected", "call-held");	
	clientId.classList.add("call-connected");
	
	changeCallPresence("InACall", "Busy");
	showClock();	
	
	acceptCall.style.display = "none";	
	declineCall.style.display = "none";	
	
	searchView.style.display = "none";
	callView.style.display = "";	

	if (viewClient.id == "search" || viewClient.id == "suggest") {
		viewClient.style.display = "";
	
		if (viewClient.id == "suggest") {
			acceptSuggestion.style.display = "";	
			declineSuggestion.style.display = "";	
		}			
		return;
	}
		
	setToActiveClient();	
}

function handleConferencedCall(data) {
	clientId.classList.remove("call-connected", "call-conferenced");	
	
	if (data.participants.length > 1) {
		clientId.classList.add("call-conferenced");
	} else {
		clientId.classList.add("call-connected");	
	}
}

function handleSharedActiveCall(data) {
	console.debug("handleSharedActiveCall", data);
	
	if (getSetting("cas_enable_shared_active_call_notification")) {
		let callerName = data.request.features.callerName;
		if (callerName.startsWith("+")) callerName = telCache[callerName]?.displayName
		if (!callerName) callerName = data.request.features.callerName;
		
		data.request.features.callerName = callerName;
		data.contextMessage = "Client " + callerName + " started a call with FA " + data.request.features.calledName;
		chrome.runtime.sendMessage({action: "show_shared_active_call",  data});		
	}
}

function handleViewClientAction(target) {
	console.debug("handleViewClientAction", target);
	
	if (target.id == "view" || target.id == "suggest") {
		setToActiveClient();
		callView.style.display = "";
		searchView.style.display = "none";		
	}
	else
		
	if (target.id == "search") {
		doSearch();
	}
}

function doSearch() {
	resizeWindowSearch();		
	callView.style.display = "none";
	searchView.style.display = "";	
	
	const picker = document.querySelector("#search-people-picker");
	
	picker.addEventListener("selectionChanged", (evt) => {
		console.log("selected", evt.detail);
		
		if (evt.detail.length > 0) {
			selectedContact = evt.detail[0];	
			clientId.personDetails = evt.detail[0];			
			updateCache();
			setToActiveClient();	
		}			
	})
}

function getClient(data) {
	console.debug("getClient", data, telCache);	

	if (data.label == "Unknown Caller") {
		const stored = telCache[data.callerId];	
		
		if (stored?.displayName) {
			viewClient.innerHTML = "Suggested Match";
			viewClient.id = "suggest";
			viewClient.style.display = "";									
			return stored;
		}
 
		viewClient.innerHTML = "Search";
		viewClient.id = "search";				
		return {displayName: "No Match", scoredEmailAddresses: [{address: ""}]};
	}
	
	viewClient.innerHTML = "View Client";	
	viewClient.id = "view";	
	return {displayName: data.label, scoredEmailAddresses: [{address: ""}]};
}

function urlParam(name) {
	var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	if (!results) { return undefined; }
	return unescape(results[1] || undefined);
};

function resizeWindowFull() {
	resizeTo(600, 900);
	moveTo(1400, 100);
}

function resizeWindowSearch() {
	resizeTo(500, 700);
	moveTo(1400, 300);
}

function resizeWindowMinimised() {
	resizeTo(350, 200);
	moveTo(1550, 840);		
}

function setToActiveClient() {
	clientId.showPresence = false;	
	clientId.fallbackDetails.displayName = selectedContact.displayName;	
	clientId.showPresence = true;
	callView.style.display = "";	

	viewClient.id = "active";
	viewClient.style.display = "none";
	
	searchView.style.display = "none";		
	acceptSuggestion.style.display = "none";	
	declineSuggestion.style.display = "none";
	endCall.style.display = "";	
	resizeWindowFull();	
}

function getDestination(person) {
	console.debug("getDestination", person);
	
	let destination = null;

	if (person.personType.subclass == "OrganizationUser") {
		destination = person.id;
		
	} else if (person.phones.length > 0) {
		destination = person.phones[0].number.replaceAll(" ", ""); // TODO pick list
	}

	return destination;
}

function updateCache() {
	telCache[currentCli] = selectedContact;
	localStorage.setItem("cas.telephone.cache", JSON.stringify(telCache));
}

function doFilter(filterList) {
	const filter = filterList.value.toLowerCase();
	const panels = document.querySelectorAll('.ms-PeoplePicker-result');	

	panels.forEach((panel) => {
		panel.style.display = "block";
		const name = panel.querySelector('.ms-Persona-primaryText').innerHTML.toLowerCase();
		console.debug("oninput", filter, name);				
		
		if (filter.length > 0 && name.indexOf(filter) == -1) {
			panel.style.display = "none";
		}				
	});		
}

function getInitials(nickname) {
	if (!nickname) nickname = "Anonymous";
	var initials = nickname.substring(nickname, 0)

	var first, last, pos = nickname.indexOf("@");
	if (pos > 0) nickname = nickname.substring(0, 1);

	// try to split nickname into words at different symbols with preference
	let words = nickname.split(/[, ]/); // "John W. Doe" -> "John "W." "Doe"  or  "Doe,John W." -> "Doe" "John" "W."
	if (words.length == 1) words = nickname.split("."); // "John.Doe" -> "John" "Doe"  or  "John.W.Doe" -> "John" "W" "Doe"
	if (words.length == 1) words = nickname.split("-"); // "John-Doe" -> "John" "Doe"  or  "John-W-Doe" -> "John" "W" "Doe"

	if (words && words[0] && words.first != '') {
		const firstInitial = words[0][0]; // first letter of first word
		var lastInitial = null; // first letter of last word, if any

		const lastWordIdx = words.length - 1; // index of last word
		if (lastWordIdx > 0 && words[lastWordIdx] && words[lastWordIdx] != '')
		{
			lastInitial = words[lastWordIdx][0]; // first letter of last word
		}

		// if nickname consist of more than one words, compose the initials as two letter
		initials = firstInitial;
		
		if (lastInitial) {
			// if any comma is in the nickname, treat it to have the lastname in front, i.e. compose reversed
			initials = nickname.indexOf(",") == -1 ? firstInitial + lastInitial : lastInitial + firstInitial;
		}
	}

	return initials.toUpperCase();
}

function hideClock() {
	document.getElementById("clocktext").style.display = "none";
}

function showClock() {
	const textElem = document.getElementById("clocktext");
	textElem.style.display = "";
	const clockTrackJoins = Date.now();

	function updateClock() {
		let clockStr = formatTimeSpan((Date.now() - clockTrackJoins) / 1000);
		textElem.textContent = clockStr;
		
		if (document.getElementById("clocktext").style.display != "none") {
			setTimeout(updateClock, 1000);
		}
	}

	updateClock();
}

function formatTimeSpan(totalSeconds) {
	const secs = ('00' + parseInt(totalSeconds % 60, 10)).slice(-2);
	const mins = ('00' + parseInt((totalSeconds / 60) % 60, 10)).slice(-2);
	const hrs = ('00' + parseInt((totalSeconds / 3600) % 24, 10)).slice(-2);
	return `${hrs}:${mins}:${secs}`;
}

async function getMeetingLink(meetId) {	
	const resp1 = await fetch(getSetting("cas_server_url") + "/teams/api/openlink/shared/meeting/" + meetId, {method: "GET", headers: {authorization: getSetting("cas_server_token")}});	
	const line = await resp1.json();	
	return line.joinWebUrl;	
}

async function sendEmail(email, subject, body) {
	console.log("sendEmail", email, subject, body);
	await fetch(getSetting("cas_server_url") + `/teams/api/openlink/email/${subject}/${email}`, {method: "POST", headers: {authorization: getSetting("cas_server_token")}, body})	
}

async function sendWhatsAppInvite(data, meetingLink) {
	const telephoneNo = data.callerId;	
	const body = meetingLink.replace("https://teams.microsoft.com/l/meetup-join/", "");
	console.log("sendWhatsAppInvite", body, telephoneNo, data, meetingLink);	
	await fetch(getSetting("cas_server_url") + "/teams/api/openlink/workflow/whatsapp/" + telephoneNo + "/join_teams_meeting", {method: "POST", headers: {authorization: getSetting("cas_server_token")}, body});	
}

async function joinMeeting(key, emails, name, data) {	
	// don't current hangup. wiat for use to end call
	//chrome.runtime.sendMessage({action: "hangup_call", id: callId});
	
	const address = await getMeetingLink(key);	
	const action = "join_meeting";			
	const locator = {meetingLink: btoa(address)};
	const token = getSetting('cas_access_token');
	const userId = { microsoftTeamsUserId: getSetting('cas_endpoint_address')};	
	const displayName = getSetting('cas_endpoint_name');	
	chrome.runtime.sendMessage({key, action,  data: {locator, displayName, userId, token}});

	const body = `Hi ${name},\n\nPlease join meeting at ${address}\n\n${getSetting('cas_endpoint_name')}`;
	const subject = "Online Meeting with " + getSetting('cas_endpoint_name')
		
	for (let email of emails) {
		sendEmail(email.address, subject, body);
	}
	
	sendWhatsAppInvite(data, address);	
}

