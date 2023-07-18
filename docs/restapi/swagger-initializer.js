window.onload = function() {
  //<editor-fold desc="Changeable Configuration Block">

  // the following lines will be replaced by docker/configurator, when it runs in a docker-container
  window.ui = SwaggerUIBundle({
    url: location.host.indexOf("comitfs.github.io") > -1 ? "https://do.comitfs.com/teams/api/swagger.json" : location.protocol + "//" + location.host + "/teams/api/swagger.json",
    dom_id: '#swagger-ui',
    deepLinking: true,
    presets: [
      SwaggerUIBundle.presets.apis,
      SwaggerUIStandalonePreset
    ],
    plugins: [
      SwaggerUIBundle.plugins.DownloadUrl
    ],
    onComplete: function() {
		const urlify = (text) => {
			var urlRegex = /(https?:\/\/[^\s]+)/g;
			return text.replace(urlRegex, '<a target="_blank" href="$1">$1</a>');
		}

		const urlParam = ()	=> {
			let value = null;
			let results = new RegExp('[\\?&]t=([^&#]*)').exec(window.location.href);
			
			if (results) {
				value = unescape(results[1] || undefined);
			}
			
			if (!value) {
				value = sessionStorage.getItem("cas.workflow.config." + name);
			}	

			if (!value) {
				value = prompt("Enter User Access Token");
			}			
				
			if (value) {
				sessionStorage.setItem("cas.workflow.config." + name, value);
			}

			console.debug("urlParam get token", value);				
			return value;
		}		
		
		let token = urlParam();
		
		if (token) {
			window.ui.preauthorizeApiKey("authorization", token);	
			const url = location.host.indexOf("comitfs.github.io") > -1 ? "https://localhost:7443" : location.protocol + "//" + location.host
			const source = new EventSource(url + "/teams/web-sse?token=" + token);
			
			source.onerror = event => {
				console.debug("EventSource - onError", event);				
			};

			source.addEventListener('onNotify', async event => {
				const msg = JSON.parse(event.data);
				document.getElementById("status").innerHTML = urlify(msg.text);
				console.debug("EventSource - onMessage", msg);
			});
			
			source.addEventListener('onConnect', async event => {
				const msg = JSON.parse(event.data);				
				document.getElementById("status").innerHTML = "CAS User " + msg.username + " (" + msg.name + ") Signed In";				
				console.debug("EventSource - onConnect");		
			});	
		}			
    },	
    layout: "StandaloneLayout"
  });

  //</editor-fold>
};
