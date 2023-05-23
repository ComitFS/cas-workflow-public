let dialogComponent;

window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);	

	if (microsoftTeams in window) {
		microsoftTeams.initialize();
		microsoftTeams.appInitialization.notifyAppLoaded();

		microsoftTeams.getContext(async context => {
			microsoftTeams.appInitialization.notifySuccess();	
			console.log("cas workflow contacts logged in user", context.userPrincipalName, context);
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
	const authorization = urlParam("t");
	const url = urlParam("u") + "/teams/api/openlink/workflow/colleagues";
			
	const response =  await fetch(url, {method: "GET", headers: {authorization}});	
	const colleagues = await response.json();
	console.debug("onload", colleagues);
		
	if (colleagues && colleagues.length > 0) 
	{
		for (let colleague of colleagues) {
			let peers =  "";
			
			for (let peer of colleague.peers) {
				const phoneNo = formatPhoneNumber(peer.teams_phone);
				
				if (phoneNo && phoneNo.trim().length > 0) {
					peers += `
					  <div class="ms-PeoplePicker-result" tabindex="1">
						<div class="ms-Persona ms-Persona--sm">
						  <div class="ms-Persona-imageArea">
							<div class="ms-Persona-initials ms-Persona-initials--blue">${getInitials(peer.name)}</div>
						  </div>
						  <div class="ms-Persona-presence">
						  </div>
						  <div class="ms-Persona-details">
							<div class="ms-Persona-primaryText"><input data-group="${colleague.group}" class="${colleague.group} cas-user-selection" id="${peer.username}" type="checkbox">${peer.name}</div>
							<div class="ms-Persona-secondaryText">${phoneNo}</div>					
						  </div>
						</div>
						<button class="ms-PeoplePicker-resultAction">
						  <i class="ms-Icon ms-Icon--Clear"></i>
						</button>
					  </div>			
					`;	
				}
			}				

			const ele = document.createElement("div");
			document.body.appendChild(ele);
				
			ele.outerHTML = `
			<details id="${colleague.group}">
				<summary><span><input class="cas-group-selection" id="${colleague.group}" type="checkbox">${colleague.group} (${colleague.description})</span></summary>			
				<div>${peers}</div>				
			</details>						
						`			
		}

		const personaElements = document.querySelectorAll(".ms-Persona");
		
		for (let personaElement of personaElements) {
			new fabric.Persona(personaElement);
		}	

		const checkElements = document.querySelectorAll(".cas-group-selection");
		
		for (let checkElement of checkElements) {
			checkElement.addEventListener("click", function() {
				console.debug("group selected", this.id, this.checked);				
				const selected = document.querySelectorAll("." + this.id);				
				
				for (let item of selected) {	
					item.checked = this.checked;
				}				
			});
		}

		const selectionButton = document.querySelector("#selectionButton");
		
		selectionButton.addEventListener("click", async (ev) => {
			const selectedElements = document.querySelectorAll(".ms-Persona-secondaryText");
			const authorization = urlParam("t");			
			
			for (let selectedElement of selectedElements) {
				const checkedElement = selectedElement.parentNode.querySelector("input");
				
				if (checkedElement.checked) {
					const telNumbers = selectedElement.innerHTML.trim();
					
					for (let telNumber of telNumbers.split(" ")) {
						const url = urlParam("u") + "/teams/api/openlink/workflow/joincall/" + telNumber;						
						const response =  await fetch(url, {method: "POST", headers: {authorization}});
						
						const resultAction = selectedElement.parentNode.parentNode.parentNode.querySelector(".ms-PeoplePicker-resultAction .ms-Icon");
						console.debug("Adding participant ", telNumber, response.ok, resultAction);
						resultAction.classList.replace("ms-Icon--Clear", response.ok ? "ms-Icon--Accept" : "ms-Icon--Error");
						resultAction.style = response.ok ? "color: green;" : "color: red;";		
					}
				}					
			}			
		});		
	}
};

function formatPhoneNumber(phone) {
	let phoneNumber = "";
	
	if (phone) {
		const numberObjEvt = libphonenumber.parsePhoneNumber(phone, "US");				

		if (numberObjEvt.isValid()) {				
			phoneNumber = numberObjEvt.format('E.164') + " ";		
			const formattedPhoneNumber = numberObjEvt.format('IDD', {fromCountry: 'GB'});				
		}
	}		
	return phoneNumber;
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

