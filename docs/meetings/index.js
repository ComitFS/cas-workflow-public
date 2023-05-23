let callId = sessionStorage.getItem("callId");
let dialogComponent;

window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);
	
	if (microsoftTeams in window) {
		microsoftTeams.initialize();
		microsoftTeams.appInitialization.notifyAppLoaded();

		microsoftTeams.getContext(async context => {
			microsoftTeams.appInitialization.notifySuccess();	
			console.log("cas workflow meetings logged in user", context.userPrincipalName, context);
			setupUI();
		});

		microsoftTeams.registerOnThemeChangeHandler(function (theme) {
			console.log("change theme", theme);
		});	
	} else {
		setupUI();
	}
});

function actionHandler(event) {
	const url = document.querySelector("#server_url").value;
	const token = document.querySelector("#access_token").value;	
	
	if (url && url.length > 0 && token && token.length > 0) {
		localStorage.setItem("cas.workflow.config.u", url);
		localStorage.setItem("cas.workflow.config.t", token);
		
		console.debug("actionHandler", url, token);
		location.reload();
	}
}

function setupUI() {
	const refreshCalendar = document.querySelector('#refresh');

	refreshCalendar.addEventListener('click', async () => {
		location.reload();
	});
	
	const settingsButton = document.querySelector('#settings');

	settingsButton.addEventListener('click', async () => {
		openDialog();
	});	
	
	setupDialog();	
}

function openDialog() {
	const url = localStorage.getItem("cas.workflow.config.u");
	const token = localStorage.getItem("cas.workflow.config.t");
	
	document.querySelector("#server_url").value = url ? url : "";
	document.querySelector("#access_token").value = token ? token : "";	

	dialogComponent.open();	
}

function setupDialog() {
    const example = document.querySelector(".docs-DialogExample-close");
    const dialog = example.querySelector(".ms-Dialog");

    const textFieldsElements = example.querySelectorAll(".ms-TextField");
    const actionButtonElements = example.querySelectorAll(".ms-Dialog-action");

    dialogComponent = new fabric.Dialog(dialog);

    for (let textFieldsElement of textFieldsElements) {
      new fabric.TextField(textFieldsElement);
    }

    for (let actionButtonElement of actionButtonElements) {
      new fabric.Button(actionButtonElement, actionHandler);
    }

	if (!urlParam("t") || !urlParam("u")) {
		openDialog();
	} else {
		setupApp();
	}
}

async function setupApp()	{	
	const hangUp = document.querySelector('#terminate');

	hangUp.addEventListener('click', async () => {
		if (callId) {		
			const authorization = urlParam("t");	
			const url = urlParam("u") + "/teams/api/openlink/workflow/meeting";	
			console.log("hangup", callId);
			
			const response = await fetch(url, {method: "DELETE", headers: {authorization}, body: callId});
		}			
	});
	
	const calendarEl = document.querySelector('#full_calendar');
	
	const config =  {
		selectable: true,				
		initialView: 'timeGridDay',
		headerToolbar: {
			left: 'prev,next today',
			center: 'title',
			right: 'dayGridMonth,timeGridWeek,timeGridDay'
		},
		dateClick: function(info) {
			console.debug("dateClick", info)
		},
		select: function(info) {
			console.debug("select", info)
		},				
		events: []
	};		

	const authorization = urlParam("t");
	const url = urlParam("u") + "/teams/api/openlink/workflow/meetings";
			
	const response =  await fetch(url, {method: "GET", headers: {authorization}});	
	const meetings = await response.json();
	console.debug("setupCalendarView", calendarEl, config, meetings);
	
	if (meetings && meetings.length > 0) {

		for (let meeting of meetings) {
			const title =  meeting.subject;
			const start = meeting.start;			
			const url = `javascript:joinMeeting('${meeting.url}', '${title}', '${start}')`;
			
			config.events.push({title, url, start});
		}		
		calendar = new FullCalendar.Calendar(calendarEl, config);
		calendar.render();
	}
}

function urlParam(name)	{
	let value = null;
	let results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	
	if (results) {
		value = unescape(results[1] || undefined);
		console.debug("urlParam get", name, value);	
	}
	
	if (!value) {
		value = localStorage.getItem("cas.workflow.config." + name);
		console.debug("urlParam get", name, value);	
	}		
		
	if (value) {
		localStorage.setItem("cas.workflow.config." + name, value);
		console.debug("urlParam set", name, value);		
	}
	return value;
}

async function joinMeeting(body, title, start) {
	const authorization = urlParam("t");	
	const url = urlParam("u") + "/teams/api/openlink/workflow/meeting";	
	console.log("joinMeeting", url, title, start);
	
	if (confirm("Are you sure you wish to join " + title)) {
		const response = await fetch(url, {method: "POST", headers: {authorization}, body});
		
		if (response.ok) {
			const json = await response.json();	
			callId = json.resourceUrl;
			alert("Joining Call " + callId);
			console.debug("joinMeeting - response", response, callId);	
			sessionStorage.setItem("callId", callId);
		} else {
			alert("Unable to join meeting");
		}
	}
}
