let callId = sessionStorage.getItem("callId");
let settingsDialog, confirmDialog, activeMeeting;

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

async function actionHandler(event) {
	console.log("actionHandler", event.target.innerHTML, event.target.id);
	
	if (event.target.id == "settingsDialogSave") {
		const url = document.querySelector("#server_url").value;
		const token = document.querySelector("#access_token").value;	
		
		if (url && url.length > 0 && token && token.length > 0) {
			localStorage.setItem("cas.workflow.config.u", url);
			localStorage.setItem("cas.workflow.config.t", token);
			
			console.debug("actionHandler", url, token);
			location.reload();
		}
	}
	else
		
	if (event.target.id == "confirmDialogProceed") {
		const status = document.getElementById("status");
		status.innerHTML = "Joining....";
		const authorization = urlParam("t");	
		const url = urlParam("u") + "/teams/api/openlink/workflow/meeting";	
	
		const response = await fetch(url, {method: "POST", headers: {authorization}, body: activeMeeting.url});
		
		if (response.ok) {
			const json = await response.json();	
			callId = json.resourceUrl;
			status.innerHTML = "Joined Meeting - " + callId;
			console.debug("joinMeeting - response", response, callId);	
			sessionStorage.setItem("callId", callId);
		} else {
			status.innerHTML = "Unable to join meeting";
		}
	}		
}

function setupUI() {
	const refreshCalendar = document.querySelector('#refresh');

	refreshCalendar.addEventListener('click', async () => {
		location.reload();
	});
	
	const settingsButton = document.querySelector('#settings');

	settingsButton.addEventListener('click', async () => {
		openSettingsDialog();
	});	
	
	setupDialog();	
}

function openSettingsDialog() {
	const url = localStorage.getItem("cas.workflow.config.u");
	const token = localStorage.getItem("cas.workflow.config.t");
	
	document.querySelector("#server_url").value = url ? url : "";
	document.querySelector("#access_token").value = token ? token : "";	

	settingsDialog.open();	
}

function setupDialog() {
    const container = document.querySelector("#dialogContainer");
    const textFieldsElements = container.querySelectorAll(".ms-TextField");
    const actionButtonElements = container.querySelectorAll(".ms-Dialog-action");

    settingsDialog = new fabric.Dialog(container.querySelector("#settingsDialog"));
	confirmDialog = new fabric.Dialog(container.querySelector("#confirmDialog"));

    for (let textFieldsElement of textFieldsElements) {
      new fabric.TextField(textFieldsElement);
    }

    for (let actionButtonElement of actionButtonElements) {
      new fabric.Button(actionButtonElement, actionHandler);
    }

	if (!urlParam("t") || !urlParam("u")) {
		openSettingsDialog();
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
			document.getElementById("status").innerHTML = "Inactive";			
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

async function joinMeeting(url, title, start) {
	console.log("joinMeeting", url, title, start);
	
	activeMeeting = {url, title, start};
	const label = title + " at " + time_ago(start);
	document.getElementById("confirmDialogDesc").innerHTML = label;
	document.getElementById("status").innerHTML = label;
	confirmDialog.open();	
}

function time_ago(time) {
	switch (typeof time) {
		case 'number':
		  break;
		case 'string':
		  time = +new Date(time);
		  break;
		case 'object':
		  if (time.constructor === Date) time = time.getTime();
		  break;
		default:
		  time = +new Date();
	}
	var time_formats = [
		[60, 'seconds', 1], // 60
		[120, '1 minute ago', '1 minute from now'], // 60*2
		[3600, 'minutes', 60], // 60*60, 60
		[7200, '1 hour ago', '1 hour from now'], // 60*60*2
		[86400, 'hours', 3600], // 60*60*24, 60*60
		[172800, 'Yesterday', 'Tomorrow'], // 60*60*24*2
		[604800, 'days', 86400], // 60*60*24*7, 60*60*24
		[1209600, 'Last week', 'Next week'], // 60*60*24*7*4*2
		[2419200, 'weeks', 604800], // 60*60*24*7*4, 60*60*24*7
		[4838400, 'Last month', 'Next month'], // 60*60*24*7*4*2
		[29030400, 'months', 2419200], // 60*60*24*7*4*12, 60*60*24*7*4
		[58060800, 'Last year', 'Next year'], // 60*60*24*7*4*12*2
		[2903040000, 'years', 29030400], // 60*60*24*7*4*12*100, 60*60*24*7*4*12
		[5806080000, 'Last century', 'Next century'], // 60*60*24*7*4*12*100*2
		[58060800000, 'centuries', 2903040000] // 60*60*24*7*4*12*100*20, 60*60*24*7*4*12*100
	];
	var seconds = (+new Date() - time) / 1000,
	token = 'ago',
	list_choice = 1;

	if (seconds == 0) {
		return 'Just now'
	}
	
	if (seconds < 0) {
		seconds = Math.abs(seconds);
		token = 'from now';
		list_choice = 2;
	}
	var i = 0,
	format;
	
	while (format = time_formats[i++]) 
	{
		if (seconds < format[0]) {
		  if (typeof format[2] == 'string')
			return format[list_choice];
		  else
			return Math.floor(seconds / format[2]) + ' ' + format[1] + ' ' + token;
		}
	}
	return time;
}
