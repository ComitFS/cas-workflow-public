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
			WebAuthnGoJS.CreateContext(JSON.stringify({RPDisplayName: "CAS Workflow", RPID: window.location.hostname, RPOrigin: window.location.origin}), (err, val) => {
				if (err) {
					handleError(err);
				}
				navigator.credentials.get({password: true}).then(function(credential) {
					console.debug("window.load credential", credential);	
					
					if (credential) {
						loginUser(credential.id, credential.password);
					} else {
						registerUser();							
					}
				}).catch(function(err){
					console.error("window.load credential error", err);	
					registerUser();					
				});	
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
		const authorization = localStorage.getItem("cas.workflow.token");		
		let host = localStorage.getItem("cas.workflow.url");

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

			console.debug("handleCredentials", username, password, configData);		
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

	function bufferDecode(value) {
	  return Uint8Array.from(atob(value), c => c.charCodeAt(0));
	}

	function bufferEncode(value) {
	  return btoa(String.fromCharCode.apply(null, new Uint8Array(value)))
		.replace(/\+/g, "-")
		.replace(/\//g, "_")
		.replace(/=/g, "");;
	}

	function loginUser(username, userStr) {
		const user = JSON.parse(userStr);
		console.debug("loginUser", user);		
		
		const loginCredRequest = (credentialRequestOptions) => {
			credentialRequestOptions.publicKey.challenge = bufferDecode(credentialRequestOptions.publicKey.challenge);
			credentialRequestOptions.publicKey.allowCredentials.forEach(function (listItem) {
			  listItem.id = bufferDecode(listItem.id)
			});

			return navigator.credentials.get({
			  publicKey: credentialRequestOptions.publicKey
			})
		}

		WebAuthnGoJS.BeginLogin(userStr, (err, data) => {
			if (err) {
				console.error("Login failed", err);				
				handleError(err);
			}

			data = JSON.parse(data);
			user.authenticationSessionData = data.authenticationSessionData;

			loginCredRequest(data.credentialRequestOptions).then((assertion) => {
				let authData = assertion.response.authenticatorData;
				let clientDataJSON = assertion.response.clientDataJSON;
				let rawId = assertion.rawId;
				let sig = assertion.response.signature;
				let userHandle = assertion.response.userHandle;

				const finishLoginObj = {
					id: assertion.id,
					rawId: bufferEncode(rawId),
					type: assertion.type,
					response: {
						authenticatorData: bufferEncode(authData),
						clientDataJSON: bufferEncode(clientDataJSON),
						signature: bufferEncode(sig),
						userHandle: bufferEncode(userHandle)
					}
				}

				const loginBodyStr = JSON.stringify(finishLoginObj);
				const authSessDataStr = JSON.stringify(user.authenticationSessionData)

				WebAuthnGoJS.FinishLogin(userStr, authSessDataStr, loginBodyStr, (err, result) => {
					console.debug("Login result", username, err, result);
					
					if (err) {
						handleError(err);
					}						
					setCredentials(username, userStr);
				});
			}).catch((err) => {
				console.error("Login failed", err);
				handleError(err);				
			});
	  });
	}	

	function urlParam(name)	{
		var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
		if (!results) { return undefined; }
		return unescape(results[1] || undefined);
	}

	function registerUser() {	
		const username = urlParam("i")
		const token = urlParam("t");
		const url = urlParam("u");
				
		if (username && token && url) {
			localStorage.setItem("cas.workflow.token", token);
			localStorage.setItem("cas.workflow.url", url);		
			createUserCredentials(username, token);
		}
	}	
	
	function createUserCredentials(username, token) 
	{
		if (!username || username.trim() === "" || !token || token.trim() === "") {
			handleError("bad username or token");
		}
		
		const displayName = username;
		username = username.toLocaleLowerCase().replaceAll(" ", "");		

		const createPromiseFunc = (credentialCreationOptions) => 
		{
			credentialCreationOptions.publicKey.challenge = bufferDecode(credentialCreationOptions.publicKey.challenge);
			credentialCreationOptions.publicKey.user.id = bufferDecode(credentialCreationOptions.publicKey.user.id);
			
			if (credentialCreationOptions.publicKey.excludeCredentials) 
			{
			  for (var i = 0; i < credentialCreationOptions.publicKey.excludeCredentials.length; i++) {
				credentialCreationOptions.publicKey.excludeCredentials[i].id = bufferDecode(credentialCreationOptions.publicKey.excludeCredentials[i].id);
			  }
			}

			return navigator.credentials.create({
			  publicKey: credentialCreationOptions.publicKey
			})
		}

		const user = {
			id: Math.floor(Math.random() * 1000000000),
			name: username,
			displayName: displayName,
			credentials: [],
		};
  
		const userStr = JSON.stringify(user);

		WebAuthnGoJS.BeginRegistration(userStr, (err, data) => 
		{
			if (err) {
				console.error("Registration failed", err);				
				handleError(err);
			}
			
			data = JSON.parse(data);
			user.registrationSessionData = data.registrationSessionData;

			createPromiseFunc(data.credentialCreationOptions).then((credential) => {
				let attestationObject = credential.response.attestationObject;
				let clientDataJSON = credential.response.clientDataJSON;
				let rawId = credential.rawId;

				const registrationBody = {
					id: credential.id,
					rawId: bufferEncode(rawId),
					type: credential.type,
					response: {
					  attestationObject: bufferEncode(attestationObject),
					  clientDataJSON: bufferEncode(clientDataJSON),
					},
				};

				// Stringify
				const regBodyStr = JSON.stringify(registrationBody);
				const sessDataStr = JSON.stringify(user.registrationSessionData)

				WebAuthnGoJS.FinishRegistration(userStr, sessDataStr, regBodyStr, (err, result) => 
				{
					if (err) {
						console.error("Registration failed", err);				
						handleError(err);
					}
					
					const credential = JSON.parse(result);
					credential.github_token = token;
					user.credentials.push(credential);						
					registerCredential(username, JSON.stringify(user));

				});
				
			}).catch((err) => {
				console.error("Registration failed", err);
				handleError(err);
			});
		})
	}
	
	async function hashCode(target){
	   var buffer = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(target));
	   var chars = Array.prototype.map.call(new Uint8Array(buffer), ch => String.fromCharCode(ch)).join('');
	   return btoa(chars);
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
			chrome.runtime.sendMessage('ahmnkjfekoeoekkbgmpbgcanjiambfhc', configData);
		} else {
			alert("You are not authorizsed to do this");
		}		
	}

    return api;

}(cas_workflow_api || {}));