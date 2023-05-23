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
	const url = urlParam("u") + "/teams/api/openlink/workflow/people";
			
	const response =  await fetch(url, {method: "GET", headers: {authorization}});	
	const people = await response.json();
	console.debug("onload", people);
		
	if (people && people.length > 0) {
		const resultGroup = document.querySelector(".ms-PeoplePicker-resultGroup");

		for (let person of people) {
			let accum = "";
			
			for (let phone of person.phones) {
				const numberObjEvt = libphonenumber.parsePhoneNumber(phone.number, "US");				

				if (numberObjEvt.isValid()) {				
					const phoneNumber = numberObjEvt.format('E.164');		
					const formattedPhoneNumber = numberObjEvt.format('IDD', {fromCountry: 'GB'});
					accum += (phoneNumber + " ");					
				}
			}

			if (accum.length > 0) {
				const ele = document.createElement("div");
				resultGroup.appendChild(ele);
				ele.outerHTML = `
				  <div class="ms-PeoplePicker-result" tabindex="1">
					<div class="ms-Persona ms-Persona--sm">
					  <div class="ms-Persona-imageArea">
						<div class="ms-Persona-initials ms-Persona-initials--blue">${getInitials(person.displayName)}</div>
					  </div>
					  <div class="ms-Persona-presence">
					  </div>
					  <div class="ms-Persona-details">
						<div class="ms-Persona-primaryText">${person.displayName}</div>
						<div class="ms-Persona-secondaryText">${accum}</div>
					  </div>
					</div>
					<button class="ms-PeoplePicker-resultAction">
					  <i class="ms-Icon ms-Icon--Clear"></i>
					</button>
				  </div>			
				`;	
			}				
		}
		
		const peoplePickerElements = document.querySelectorAll(".ms-PeoplePicker");
		
		for (let peoplePickerElement of peoplePickerElements) {
			new fabric.PeoplePicker(peoplePickerElement);
		}
		
		const selectionButton = document.querySelector("#selectionButton");
		
		selectionButton.addEventListener("click", async (ev) => {
			const selectedElements = document.querySelectorAll(".ms-Persona-secondaryText");
			const authorization = urlParam("t");			
			
			for (let selectedElement of selectedElements) {
				const telNumbers = selectedElement.innerHTML.trim();
				
				for (let telNumber of telNumbers.split(" ")) {
					const url = urlParam("u") + "/teams/api/openlink/workflow/joincall/" + telNumber;						
					const response =  await fetch(url, {method: "POST", headers: {authorization}});
					
					const resultAction = selectedElement.parentNode.parentNode.parentNode.querySelector(".ms-PeoplePicker-resultAction .ms-Icon");
					console.debug("Adding participant ", telNumber, response.ok, resultAction);
					resultAction.classList.replace("ms-Icon--Cancel", response.ok ? "ms-Icon--Accept" : "ms-Icon--Error");
					resultAction.style = response.ok ? "color: green;" : "color: red;";					
				}					
			}			
		});
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

function newElement(parent, el, id, html, class1, class2, class3, class4, class5) {
	const ele = document.createElement(el);
	if (id) ele.id = id;
	if (html) ele.innerHTML = html;
	if (class1) ele.classList.add(class1);
	if (class2) ele.classList.add(class2);
	if (class3) ele.classList.add(class3);
	if (class4) ele.classList.add(class4);
	if (class5) ele.classList.add(class5);	
	parent.appendChild(ele);
	return ele;
}

