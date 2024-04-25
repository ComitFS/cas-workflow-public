// -------------------------------------------------------
//
//  Functions
//
// -------------------------------------------------------


function loadJS(name) {
	console.debug("loadJS", name);
	var head  = document.getElementsByTagName('head')[0];
	var s1 = document.createElement('script');
	s1.src = name;
	s1.async = false;
	head.appendChild(s1);
}
	
function setSetting(name, value) {
    console.debug("setSetting", name, value);
    window.localStorage["store.settings." + name] = JSON.stringify(value);
}

function setDefaultSetting(name, defaultValue) {
    console.debug("setDefaultSetting", name, defaultValue, window.localStorage["store.settings." + name]);

    if (!window.localStorage["store.settings." + name] && window.localStorage["store.settings." + name] != false)
    {
        if (defaultValue) window.localStorage["store.settings." + name] = JSON.stringify(defaultValue);
    }
}

function getSetting(name, defaultValue) {
     var value = defaultValue ? defaultValue : null;

    if (window.localStorage["store.settings." + name] && window.localStorage["store.settings." + name] != "undefined") {
        value = JSON.parse(window.localStorage["store.settings." + name]);
    }
    return value;
}

function removeSetting(name) {
    localStorage.removeItem("store.settings." + name);
}

function getPassword(password) {
    if (!password || password == "") return null;
    if (password.startsWith("token-")) return atob(password.substring(6));

    window.localStorage["store.settings.password"] = JSON.stringify("token-" + btoa(password));
    return password;
}