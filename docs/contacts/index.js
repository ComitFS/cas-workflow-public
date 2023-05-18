
window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);	

	const authorization = urlParam("t");
	const url = urlParam("u") + "/teams/api/openlink/workflow/contacts";
			
	const response =  await fetch(url, {method: "GET", headers: {authorization}});	
	const contacts = await response.json();
	console.debug("onload", contacts);
		
	if (contacts && contacts.length > 0) {
		const resultGroup = document.querySelector(".ms-PeoplePicker-resultGroup");

		for (let person of contacts) {
			let accum = formatPhoneNumber(person.mobilePhone);
			
			for (let phone of person.businessPhones) {
				accum += formatPhoneNumber(phone);
			}
			
			for (let phone of person.homePhones) {
				accum += formatPhoneNumber(phone);
			}			

			if (accum.length > 0) {
				const ele = document.createElement("div");
				resultGroup.appendChild(ele);
				ele.outerHTML = `
				  <div class="ms-PeoplePicker-result" tabindex="1">
					<div class="ms-Persona ms-Persona">
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
					const url = urlParam("u") + "/teams/api/openlink/workflow/call/" + telNumber;						
					const response =  await fetch(url, {method: "POST", headers: {authorization}});
					
					const resultAction = selectedElement.parentNode.parentNode.parentNode.querySelector(".ms-PeoplePicker-resultAction .ms-Icon");
					console.debug("Adding participant ", telNumber, response.ok, resultAction);
					resultAction.classList.replace("ms-Icon--Cancel", response.ok ? "ms-Icon--Accept" : "ms-Icon--Error");
					resultAction.style = response.ok ? "color: green;" : "color: red;";					
				}					
			}			
		});
	}
});

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
	var value = localStorage.getItem("cas.workflow.meeting." + name);
	if (value) return value;
	
	var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	if (!results) { return undefined; }
	value = unescape(results[1] || undefined);
	localStorage.setItem("cas.workflow.meeting." + name, value);
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

