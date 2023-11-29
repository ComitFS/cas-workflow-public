let configData;
	
var cas_workflow_api = (function(api)
{
    window.addEventListener("unload", function()
    {
        console.debug("cas_workflow_api addListener unload");
    });

    window.addEventListener("load", function()  {
		console.debug("window.load", window.location.hostname, window.location.origin);

		const username = sessionStorage.getItem("cas.workflow.user");
		const password = sessionStorage.getItem("cas.workflow.password");	

		if (!username || !password) {
			registerUser();
			
			navigator.credentials.get({password: true}).then(function(credential) {
				console.debug("window.load credential", credential);	
				
				if (credential) {
					setCredentials(credential.id, credential.password);
				} else {
					registerUser();							
				}
			}).catch(function(err){
				console.error("window.load credential error", err);	
				registerUser();					
			});				
		} else {
			handleCredentials(username, password);
		}		
    });
	
	function handleError(err) {
		console.error(err);
	}

	function setCredentials(username, password) {
		sessionStorage.setItem("cas.workflow.user", username);
		sessionStorage.setItem("cas.workflow.password", password);			
		location.reload();
	}
	
	async function handleCredentials(username, password) {		
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

			cas_workflow_api.configure();			
		}			
	}

    function loadJS(name) {
		console.debug("loadJS", name);
        var s1 = document.createElement('script');
        s1.src = name;
        s1.async = false;
        document.body.appendChild(s1);
    }

    function loadCSS(name) {
        var head  = document.getElementsByTagName('head')[0];
        var link  = document.createElement('link');
        link.rel  = 'stylesheet';
        link.type = 'text/css';
        link.href = name;
        head.appendChild(link);
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

	function registerUser() {	
		let username = urlParam("i")
		let token = urlParam("t");
		let url = urlParam("u");
				
		if (username && token && url) {	
			registerCredential(username, token);
			
		} else {
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
			
			username = localStorage.getItem("cas.workflow.config.i");
			url = localStorage.getItem("cas.workflow.config.u");			
			token = localStorage.getItem("cas.workflow.config.t");
			
			document.querySelector("#user_name").value = username ? username : "";
			document.querySelector("#server_url").value = url ? url : "";			
			document.querySelector("#access_token").value = token ? token : "";	

			dialogComponent.open();				
		}
	}	

	function actionHandler(event) {
		const url = document.querySelector("#server_url").value;
		const token = document.querySelector("#access_token").value;
		const username = document.querySelector("#user_name").value;
		
		if (url && url.length > 0 && token && token.length > 0 && username && username.length > 0) {
			localStorage.setItem("cas.workflow.config.u", url);
			localStorage.setItem("cas.workflow.config.t", token);
			localStorage.setItem("cas.workflow.config.i", username);
			
			console.debug("actionHandler", url, token, username);
			registerCredential(username, token);
		}
	}
	
	function registerCredential(id, pass) {
		navigator.credentials.create({password: {id: id, password: pass}}).then(function(credential)
		{
			console.debug("registerCredential", credential);
		
			if (credential) {
				navigator.credentials.store(credential).then(function()
				{
					console.log("registerCredential - storeCredentials stored");				
					setCredentials(id, pass);				

				}).catch(function (err) {
					console.error("registerCredential - storeCredentials error", err);
				});
			}

		}).catch(function (err) {
			console.error("registerCredential - storeCredentials error", err);		
		});			
	}	

    //-------------------------------------------------------
    //
    //  UI
    //
    //-------------------------------------------------------	
		

    //-------------------------------------------------------
    //
    //  External
    //
    //-------------------------------------------------------
	
	api.configure = function() {
		if (configData) {
			chrome.runtime.sendMessage('ifohdfipnpbkalbeaefgecjkmfckchkd', configData);
			chrome.runtime.sendMessage('ahmnkjfekoeoekkbgmpbgcanjiambfhc', configData);			
		} else {
			alert("You are not authorizsed to do this");
		}		
	}

    return api;

}(cas_workflow_api || {}));