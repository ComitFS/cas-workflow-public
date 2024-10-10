let joinWebUrl;

window.addEventListener("unload", () => {
	console.debug("unload");
});

window.addEventListener("load", async () =>  {
	const origin = JSON.parse(localStorage.getItem("configuration.cas_server_url"));
	const authorization = JSON.parse(localStorage.getItem("configuration.cas_server_token"));
 
	console.debug("window.load", window.location.hostname, origin, authorization);
	
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
		console.log("setup", origin, authorization);
							
		document.querySelector("button").addEventListener("click", async (evt) => {	
			const url2 = origin + "/plugins/casapi/v1/companion/meeting/client/jjgartland?subject=Call%20Dect%20Phone&destination=%2B441634251467";			
			const response3 = await fetch(url2, {method: "POST", headers: {authorization}});
			joinWebUrl = await response3.text();
		
			microsoftTeams.executeDeepLink(joinWebUrl);
		})
	}	
});	

