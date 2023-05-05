
window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);	

	const authorization = urlParam("t");
	const url = urlParam("u") + "/teams/api/openlink/workflow/people";
			
	const response =  await fetch(url, {method: "GET", headers: {authorization}});	
	const people = await response.json();
	console.debug("onload", people);
		
	if (people && people.length > 0) {
		const resultGroup = document.querySelector(".ms-PeoplePicker-resultGroup");

		for (let person of people) {
			if (person.phones.length == 0) continue;

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
					<div class="ms-Persona-secondaryText">${person.phones.reduce((accum, b) => {return accum + " " + b.number}, "")}</div>
				  </div>
				</div>
				<button class="ms-PeoplePicker-resultAction">
				  <i class="ms-Icon ms-Icon--Clear"></i>
				</button>
			  </div>			
			`;			
		}
		
		const peoplePickerElements = document.querySelectorAll(".ms-PeoplePicker");
		
		for (let peoplePickerElement of peoplePickerElements) {
			new fabric.PeoplePicker(peoplePickerElement);
		}
	}
});

function urlParam(name)	{
	var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
	if (!results) { return undefined; }
	return unescape(results[1] || undefined);
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

