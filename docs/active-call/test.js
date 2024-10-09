window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	console.debug("window.load", window.location.hostname, window.location.origin);
	
	const urlParam = (name) => {
		var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
		if (!results) { return undefined; }
		return unescape(results[1] || undefined);
	}; 	
	
	if (!!microsoftTeams) {
		microsoftTeams.initialize();
		microsoftTeams.appInitialization.notifyAppLoaded();

		microsoftTeams.getContext(async context => {
			microsoftTeams.appInitialization.notifySuccess();
			console.log("cas companion logged in user", context, context.subEntityId);
			setup();
		});

		microsoftTeams.registerOnThemeChangeHandler(function (theme) {
			console.log("change theme", theme);
		});	
	}

	function setup() {
		document.querySelector("button").addEventListener("click", async (evt) => {
			let origin = urlParam("origin");	
			if (!origin) origin = "http://localhost"
			const authorization = btoa("dele:Welcome123");			
			const url2 = origin + "/plugins/casapi/v1/companion/meeting/client/jjgartland?subject=FA calling JJ Gartland&destination=+441634251467";			
			const response3 = await fetch(url2, {method: "POST", headers: {authorization}});
			const joinWebUrl = await response3.text();
			
			microsoftTeams.executeDeepLink(joinWebUrl);
		})
	}	
});	

