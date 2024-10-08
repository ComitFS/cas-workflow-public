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
			const extnId = await getExtensionId();
			console.log("cas companion logged in user", context, context.subEntityId, extnId);
			
			if (extnId) {
				location.href = "chrome-extension://" + extnId + "/cas-wealth/index.html";
			} else {
				location.href = "test.html";
			}
		});

		microsoftTeams.registerOnThemeChangeHandler(function (theme) {
			console.log("change theme", theme);
		});	
	}	
});	

async function getExtensionId() {
	let response;
	
	try {
		response = await fetch("chrome-extension://ifohdfipnpbkalbeaefgecjkmfckchkd/manifest.json");
		console.debug("getExtensionId demo", response.status);
		if (response.ok) return "ifohdfipnpbkalbeaefgecjkmfckchkd";		
		
	} catch (e) {
		console.debug("getExtensionId demo", e);
	}

	try {	
		response = await fetch("chrome-extension://ahmnkjfekoeoekkbgmpbgcanjiambfhc/manifest.json");
		console.debug("getExtensionId dev", response.status);
		if (response.ok) return "ahmnkjfekoeoekkbgmpbgcanjiambfhc";	
	} catch (e) {
		console.debug("getExtensionId dev", e);
	}
	
	return null;	
}