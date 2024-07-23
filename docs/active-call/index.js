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
			//if (context.subEntityId) config = JSON.parse(context.subEntityId);
			//config.userPrincipalName = context.userPrincipalName
			console.log("cas companion logged in user", context, context.subEntityId);
			const json = {
				type: "incoming",
				id: "1234567890",
				callerId: "+441634251467",	
				emailAddress: "dele@4ng.net"
			}
			const data = escape(JSON.stringify(json));
			location.href = "chrome-extension://ahmnkjfekoeoekkbgmpbgcanjiambfhc/cas-wealth/index.html?data=" + data;
		});

		microsoftTeams.registerOnThemeChangeHandler(function (theme) {
			console.log("change theme", theme);
		});	
	}	
});	