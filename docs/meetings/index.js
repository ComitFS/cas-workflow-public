let callId = sessionStorage.getItem("callId");

window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);

	const refreshCalendar = document.querySelector('#refresh');

	refreshCalendar.addEventListener('click', async () => {
		location.reload();
	});
	
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

		for (var meeting of meetings) {
			const title =  meeting.subject;
			const start = meeting.start;			
			const url = `javascript:joinMeeting('${meeting.url}', '${title}', '${start}')`;
			
			config.events.push({title, url, start});
		}		
		calendar = new FullCalendar.Calendar(calendarEl, config);
		calendar.render();
	}
});

function urlParam(name)	{
	var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	if (!results) { return undefined; }
	return unescape(results[1] || undefined);
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
