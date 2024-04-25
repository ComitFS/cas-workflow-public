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
		});

		microsoftTeams.registerOnThemeChangeHandler(function (theme) {
			console.log("change theme", theme);
		});	
	}
});
