let loginModal, customerData;
	
var cas_contactsi_api = (function(api)
{
	const nickColors = {};
	
    window.addEventListener("unload", () =>
    {
        console.debug("cas_contactsi_api addListener unload");
    });

    window.addEventListener("load", () =>  {
		console.debug("window.load", window.location.hostname, window.location.origin);
		
		const username = sessionStorage.getItem("cas.contacts.user");
		const password = sessionStorage.getItem("cas.contacts.password");	

		if (!username || !password) {
			WebAuthnGoJS.CreateContext(JSON.stringify({RPDisplayName: "CAS Contacts", RPID: window.location.hostname, RPOrigin: window.location.origin}), (err, val) => {
				if (err) {
					//location.href = "/";
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

	function setCredentials(username, password) {
		sessionStorage.setItem("cas.contacts.user", username);
		sessionStorage.setItem("cas.contacts.password", password);			
		location.reload();
	}
	
	async function handleCredentials(username, password) {
		const host = localStorage.getItem("cas.host");		
		console.debug("", username, password);
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
				//location.href = "/";
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
						//location.href = "/";
					}						
					setCredentials(username, userStr);
				});
			}).catch((err) => {
				console.error("Login failed", err);
				//location.href = "/";				
			});
	  });
	}	

	function urlParam(name)	{
		var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
		if (!results) { return undefined; }
		return unescape(results[1] || undefined);
	}

	function registerUser() {
		const username = urlParam("u")
		const token = urlParam("t");	
		const host = urlParam("h");		
		
		if (!username || !token) {
			alert("You do not have permissions to access this page");
		} else {
			localStorage.setItem("cas.host", host);
			createUserCredentials(username, token);
		}
	}	
	
	function createUserCredentials(username, token) 
	{
		if (!username || username.trim() === "" || !token || token.trim() === "") {
			//location.href = "/";
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
				//location.href = "/";
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
						//location.href = "/";
					}
					
					const credential = JSON.parse(result);
					credential.github_token = token;
					user.credentials.push(credential);						
					registerCredential(username, JSON.stringify(user));

				});
				
			}).catch((err) => {
				console.error("Registration failed", err);
				//location.href = "/";				
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
					console.debug("registerCredential - storeCredentials stored");				
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
	
		
    return api;

}(cas_contactsi_api || {}));