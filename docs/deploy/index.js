let configData;
	
var cas_workflow_api = (function(api)
{
    window.addEventListener("unload", function()
    {
        console.debug("cas_workflow_api addListener unload");
    });

    window.addEventListener("load", async function()  {
		console.debug("window.load", window.location.hostname, window.location.origin);

		const username =  urlParam("i");
		const password = urlParam("t");			
		const host =  urlParam("u");
		const authorization = urlParam("t");
		
		console.info("handleCredentials", username, password);		

		if (authorization && host) {
			let url = host + "/teams/api/openlink/config/global";	
			
			let response = await fetch(url, {method: "GET"});
			const config = await response.json();			
			console.info("handleCredentials config", config);
				
			url = host + "/teams/api/openlink/config/properties";	
			response = await fetch(url, {method: "GET", headers: {authorization}});
			const property = await response.json();	
			console.log("User properties", property);		

			const payload = {action: 'config', config, property, host};
			configData = JSON.stringify(payload);

			console.debug("handleCredentials", configData);	

			configure();			
		}			
	});

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
	
	function configure() {
		console.log("configure", configData);
		
		if (configData) {
			chrome.runtime.sendMessage('ifohdfipnpbkalbeaefgecjkmfckchkd', configData);
			chrome.runtime.sendMessage('ahmnkjfekoeoekkbgmpbgcanjiambfhc', configData);
			alert("Configration data sent");			
		} else {
			alert("You are not authorizsed to do this");
		}		
	}

    return api;

}(cas_workflow_api || {}));