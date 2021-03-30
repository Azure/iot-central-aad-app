import React from 'react';
import * as axios from 'axios';
import './App.css';
import * as msal from '@azure/msal-browser';

// These are the Scopes
export const Scopes = {
  Graph: 'User.Read',
  Central: 'https://apps.azureiotcentral.com/user_impersonation',
  ARM: 'https://management.azure.com/user_impersonation'
}

// This is a wrapper around the way to call MSAL to authorize against the authority for a given scope
function getAccessTokenForScope(msalInstance, scope, account) {
  const tokenRequest = {
    scopes: Array.isArray(scope) ? scope : [scope],
    forceRefresh: false,
    redirectUri: 'http://localhost:4001'
  };

  if (account) { tokenRequest.account = account };

  return new Promise((resolve, reject) => {
    msalInstance.acquireTokenSilent(tokenRequest)
      .then((res) => {
        resolve(res)
      })
      .catch((err) => {
        msalInstance.acquireTokenPopup(tokenRequest)
          .then((res) => {
            resolve(res)
          })
          .catch((err) => {
            reject(err);
          })
      })
  });
}

// This is a generic function to call an api doing auth with a bearer token
async function callAPI(instance, scope, url, account) {
  try {
    // if already authorized, this is a call to the cached token
    const at = await getAccessTokenForScope(instance, scope, account);
    const res = await axios(url, { headers: { Authorization: 'Bearer ' + at.accessToken } })
    return res;
  } catch (err) {
    alert(err);
  }
}

// This is a React helper function to make async API calls
const usePromise = ({ promiseFn }) => {
  const [loading, setLoading] = React.useState(false);
  const [data, setData] = React.useState(null);
  const [error, setError] = React.useState(null);

  const callPromise = async () => {
    setLoading(true);
    setData(null);
    setError(null);
    try {
      const res = await promiseFn();
      setData(res);
    } catch (error) {
      setError(error);
    }
    setLoading(false);
  };
  return [loading, data, error, callPromise];
};

// This is the main React component
function App() {

  // ADD Application configuration
  const [tenantId, setTenantId] = React.useState('');
  const [clientId, setClientId] = React.useState('');

  // MSAL token objects
  const [instance, setInstance] = React.useState();
  const [account, setAccount] = React.useState();
  const [central, setCentral] = React.useState();
  const [arm, setArm] = React.useState();

  // UX
  const [toggles, setToggler] = React.useState({});
  const [injectAccount, setInjectAccount] = React.useState(false);
  const [subscriptionId, setSubscriptionId] = React.useState('');

  // eslint-disable-next-line
  const [progressFetchMe, me, errorFetchMe, fetchMe] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.Graph, `https://graph.microsoft.com/v1.0/me`, injectAccount ? account.graphToken.account : null)
  });

  // eslint-disable-next-line
  const [progressFetchTemplates, templates, errorFetchTemplate, fetchTemplates] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.Central, `https://${appHost}/api/preview/deviceTemplates`, injectAccount ? account.graphToken.account : null)
  });

  // eslint-disable-next-line
  const [progressFetchDevices, devices, errorFetchDevices, fetchDevices] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.Central, `https://${appHost}/api/preview/devices`, injectAccount ? account.graphToken.account : null)
  });

  // eslint-disable-next-line
  const [progressFetchSubscription, subscriptions, errorFetchSubscription, fetchSubscription] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.ARM, `https://management.azure.com/subscriptions?api-version=2020-01-01`, injectAccount ? account.graphToken.account : null)
  });

  // eslint-disable-next-line
  const [progressFetchTenants, tenants, errorFetchTenants, fetchTenants] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.ARM, `https://management.azure.com/tenants?api-version=2020-01-01`, injectAccount ? account.graphToken.account : null)
  });

  // eslint-disable-next-line
  const [progressFetchApps, apps, errorFetchApps, fetchApps] = usePromise({
    promiseFn: () => callAPI(instance.msal, Scopes.ARM, `https://management.azure.com/subscriptions/${subscriptionId}/providers/Microsoft.IoTCentral/IoTApps?api-version=2018-09-01`, injectAccount ? account.graphToken.account : null)
  });

  // Step 1
  const updateClient = () => {
    const instance = new msal.PublicClientApplication({
      auth: {
        clientId: clientId,
        authority: 'https://login.microsoftonline.com/' + tenantId
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      }
    });
    setInstance({ msal: instance, cachedAccounts: instance.getAllAccounts() })
  }

  // Step 2
  const getGraphToken = () => {
    getAccessTokenForScope(instance.msal, Scopes.Graph, injectAccount ? account.graphToken.account : null)
      .then((res) => { setAccount({ graphToken: res }); })
      .catch((err) => { alert(JSON.stringify(err, null, 2)); })
  }

  // Step 3a
  const [appHost, setAppHost] = React.useState('');

  // Step 3b
  const getCentralToken = () => {
    getAccessTokenForScope(instance.msal, Scopes.Central, injectAccount ? account.graphToken.account : null)
      .then((res) => { setCentral({ centralAccessToken: res }); })
      .catch((err) => { alert(JSON.stringify(err, null, 2)); })
  }

  // Step 4
  const getArmToken = () => {
    getAccessTokenForScope(instance.msal, Scopes.ARM, injectAccount ? account.graphToken.account : null)
      .then((res) => { setArm({ armAccessToken: res }); })
      .catch((err) => { alert(JSON.stringify(err, null, 2)); })
  }

  // Step 1 - Sub task
  const clearStorage = () => {
    localStorage.clear();
    sessionStorage.clear();
    alert('Session and local storage cleared. Please shutdown and re-launch the browser');
  }

  const signOut = (instance) => {
    instance.logout();
  }

  function showStatus(condition, success, error) {
    return <div>{condition ? <span className='success'>{success}</span> : <span className='error'>{error}</span>}</div>
  }

  function showJSON(toggles, json, handler, ns) {
    const dom = [];
    for (const name in json) {
      const key = ns + name;
      const expand = toggles[key] ? toggles[key] : false;
      dom.push(<div className='toggle-item'><div>{name}</div><div onClick={() => { handler(key) }}>{expand ? <span>&#9660;</span> : <span>&#9650;</span>}</div>{expand ? <pre>{JSON.stringify(json[name], null, 2)}</pre> : null}</div>)
    }
    return <div className='toggle-display'>{dom}</div>;
  }

  const toggle = (key) => {
    const toggle = Object.assign({}, toggles);
    toggle[key] = toggle[key] ? !toggle[key] : true;
    setToggler(toggle)
  }

  return (<div className='App'>

    <h1>Authentication and Authorization walk-through for using Azure IoT Central REST APIs</h1>
    <p>This codebase demonstrates flows to authenticate and authorize a user to call Azure IoT Central REST APIs for data and control plane operations. It is best suited to single directory authentication models and should not be used to build multi-tenancy applications.</p>
    <h2>Prerequisites</h2>
    <ul>
      <li>A Microsoft account</li>
      <li>An Azure Subscription</li>
      <li>An Azure Active Directory with you added as admin</li>
      <li>The ability to create and admin an Azure Active Directory application using the above Subscription/Directory</li>
    </ul>
    <h4>AAD Application scopes</h4>
    <p>To learn more about setting up an AAD Application please visit <a href='https://github.com/iot-for-all/iotc-aad-setup'>this</a> repo. Ensure the following scopes are configured in the application</p>
    <ul>
      <li>Graph -'User.Read'</li>
      <li>Central - 'https://apps.azureiotcentral.com/user_impersonation'</li>
      <li>ARM - 'https://management.azure.com/user_impersonation'</li>
    </ul>

    <p>This sample code should be used for reference and pattern adoption. <b>It is not recommended for hosting/production as it exposes bearer tokens to the screen.</b> If your SPA framework is React, it is recommended to use the hooks implementation within the library. Please set the browser to always allow for pop ups from http://locahost:3000. This will ease debugging when setting up the AAD application.</p>

    {/* Step 1 */}
    <hr />
    <h2>1. Set up the MSAL instance</h2>
    <p>The MSAL framework makes doing authentication against AAD pain free by taking care of caching and refreshing all obtained tokens. This browser version includes handling all types of user interactions to capture user credentials. Visit the <a href='https://github.com/AzureAD/microsoft-authentication-library-for-js'>MSAL.js</a> github repo for details on this library. <b>This will be the only dependency required for your web application.</b></p>
    <div>
      <label>AAD Application Client ID</label><br />
      <input type='text' value={clientId} onChange={(e) => setClientId(e.target.value)} placeholder='e.g. f095895b-0529-4839-af23-58eb3aa12a54'></input>
      <br />
      <label>AAD Directory Tenant ID</label><br />
      <input type='text' value={tenantId} onChange={(e) => setTenantId(e.target.value)} placeholder='e.g. 0786922d-1889-4538-9695-dd37032b2a39'></input>
    </div>
    <br />
    <button onClick={() => { updateClient(); }}>Create MSAL instance</button>
    <br /><br />
    <label>Result(s)</label><br />
    {showStatus(instance, showJSON(toggles, instance, toggle, 'instance'), 'MSAL.js instance has not been created')}
    <br />
    <div>
      <button className='btn-sm' onClick={() => { clearStorage(); }}>Clear local/session storage data</button>
      <button disabled={!instance} className='btn-sm' onClick={() => { signOut(instance.msal); }}>Sign out if already authenticated</button>
    </div>

    {/* Step 2 */}
    <hr />
    <h2>2. Optional - Authorize against MS Graph to gain access to user's profile data. Additionally get an account</h2>
    <p>Starting authentication with MS Graph authorization will allow you to obtain an account object that can be injected into subsequent authorization calls (you can also use one of the cached accounts) Without this, each authorization call will require MSAL to use the account associate to the scope request increasing the time through the authentication cycle. The added benefit of making this call first is to get user profile information such as name and email</p>
    <button disabled={!instance} onClick={() => { getGraphToken(); }}>Auth to get access token for MS Graph APIs</button>
    <br /><br />
    <input disabled={!account} type="checkbox" value={injectAccount} onChange={(e) => setInjectAccount(e.target.checked)} /> Inject the account returned from this call into all subsequent calls for request tokens or API calls
    <br />
    <label>Result</label><br />
    {showStatus(account, showJSON(toggles, account, toggle, 'account'), 'Need to obtain MS Graph access token')}
    {showStatus(instance, '', 'Needs Step 1')}
    <br />
    <label>MS Graph APIs</label><br />
    <span>To see the full list of MS Graph REST APIs visit <a href='https://docs.microsoft.com/en-us/azure/active-directory/develop/microsoft-graph-intro'>this</a> documentation site.</span>
    <br /><br />
    <div>
      <button disabled={!account} className='btn-sm' onClick={() => { fetchMe(); }}>Fetch Me</button>
    </div>
    {progressFetchMe && !me ? <><br />{'Fetching Me'}<br /></> : me ? <><br />Me<br />{showStatus(me, showJSON(toggles, me, toggle, 'me'), '')}</> : null}

    {/* Step 3a */}
    <hr />
    <h2>3a. Authorize against Central to gain access to data plane REST APIs  </h2>
    <p><b>Skip if only control plane access is required</b></p>
    <p>This authorization is required to fetch data from your IoT Central application. If an account has not been provided with the scope request, the user may have to login or be redirected through a silent login phase.</p>
    <button disabled={!instance} onClick={() => { getCentralToken(); }}>Get access token for IoT Central REST APIs</button>
    <br /><br />
    <label>Result</label><br />
    {showStatus(central, showJSON(toggles, central, toggle, 'central'), 'Need to obtain IoT Central REST API access token')}
    {showStatus(instance, '', 'Need Step 1')}

    {/* Step 3b */}
    <hr />
    <h2>3b. Define the IoT Central application host name to call IoT Central data plane REST APIs</h2>
    <p><b>Skip if only control plane access is required</b></p>
    <p>This is required for the domain the REST calls need to use to do data plane operations. It is the given by the application name + the iotcenntral domain name e.g. myapp.azureiotcentral.com. If an account has not been provided with the scope request, the user may have to login or be redirected. The application name is also available through the applications call using the ARM APIs (see later) and is the subdomain property.</p>
    <label>Application host name</label><br />
    <div>
      <input type='text' value={appHost} onChange={(e) => setAppHost(e.target.value)} placeholder='e.g. <appname>.azureiotcentral.com'></input>
    </div>
    <br />
    <label>Result</label><br />
    {showStatus(appHost !== '' && appHost.indexOf('.com') >= 0 && appHost.indexOf('azureiotcentral') >= 0, 'Application host defined', 'No application host defined')}
    {showStatus(central, '', 'Need Step 3a')}
    <br />
    <label>APIs</label><br />
    <span>To see the full list of IoT Central REST APIs visit <a href='https://docs.microsoft.com/en-us/rest/api/iotcentral/'>this</a> documentation site.</span>
    <br /><br />
    <div>
      <button disabled={!instance || !central || appHost === ''} className='btn-sm' onClick={() => { fetchTemplates(); }}>Fetch applications templates</button>
      <button disabled={!instance || !central || appHost === ''} className='btn-sm' onClick={() => { fetchDevices(); }}>Fetch all devices from application</button>
    </div>
    {progressFetchTemplates && !templates ? <><br />{'Fetching Templates'}<br /></> : templates ? <><br />Templates<br />{showStatus(templates, showJSON(toggles, templates, toggle, 'templates'), '')}</> : null}
    {progressFetchDevices && !devices ? <><br />{'Fetching Devices'}<br /></> : devices ? <><br />Devices<br />{showStatus(devices, showJSON(toggles, devices, toggle, 'devices'), '')}</> : null}

    {/* Step 3c */}
    <hr />
    <h2>3c. Optional - Check Single Sign On with Azure IoT Central and Microsoft Outlook</h2>
    <p><b>Skip if only control plane access is required</b></p>
    <p>Because authentication has happened, any visit to another site that is using the same Scope request will mean the user will not need to sign-on again (or sign-in without needing to provide a password). Setting up MSAL.js to use localStorage enables this.</p>
    <p>Visit the following sites for SSO scenarios</p>
    <label>IoT Central application URL</label><br />
    {showStatus(central, '', 'Need Step 3a')}
    <a href={'https://' + appHost} rel='noreferrer' target='_blank'>{'https://' + appHost}</a>
    <br /><br />
    <label>Microsoft Outlook URL</label><br />
    {showStatus(account, '', 'Need Step 2')}
    <a href='https://outlook.live.com' rel='noreferrer' target='_blank'>https://outlook.live.com</a>

    {/* Step 4 */}
    <hr />
    <h2>4. Authorize against ARM to gain access to IoT Central control plane REST APIs  </h2>
    <p><b>Skip if only data plane access is required</b></p>
    <p>This authorization is required to do control plane operations for your IoT Central application. If an account has not been provided with the scope request, the user may have to login or be redirected through a silent login phase</p>
    <button disabled={!instance} onClick={() => { getArmToken(); }}>Auth to get access token for ARM REST APIs</button>
    <br /><br />
    <label>Result</label><br />
    {showStatus(arm, showJSON(toggles, arm, toggle, 'arm'), 'Need to obtain ARM REST API access token')}
    {showStatus(instance, '', 'Need Step 1')}
    <br />
    <label>ARM APIs</label><br />
    <span>To see the full list of ARM REST APIs visit <a href='https://docs.microsoft.com/en-us/rest/api/resources/'>this</a> documentation site.</span>
    <br /><br />
    <div>
      <button disabled={!arm} className='btn-sm' onClick={() => { fetchSubscription(); }}>Fetch Subscriptions</button>
      <button disabled={!arm} className='btn-sm' onClick={() => { fetchTenants(); }}>Fetch Tenants</button>
    </div>
    {progressFetchSubscription && !subscriptions ? <><br />{'Fetching Subscriptions'}<br /></> : subscriptions ? <><br />Subscriptions<br />{showStatus(subscriptions, showJSON(toggles, subscriptions, toggle, 'subs'), '')}</> : null}
    {progressFetchTenants && !tenants ? <><br />{'Fetching Tenants'}<br /></> : tenants ? <><br />Tenants<br />{showStatus(tenants, showJSON(toggles, tenants, toggle, 'tenants'), '')}</> : null}
    <br /><br />
    <label>IoT Central ARM APIs</label><br />
    <span>To see the full list of IoT Central ARM REST APIs visit <a href='https://docs.microsoft.com/en-us/azure/templates/microsoft.iotcentral/iotapps'>this</a> documentation site.</span>
    <p>For every IoT Central ARM API call, a Subscription ID is required.</p>
    <div>
      <label>Subscription ID</label><br />
      <input disabled={!arm} type='text' value={subscriptionId} onChange={(e) => setSubscriptionId(e.target.value)} placeholder='e.g. aed4abeb-3420-45ba-8ebf-5fd3b2ee891b'></input>
    </div>
    <br />
    <div>
      <button disabled={!arm || subscriptionId === ''} className='btn-sm' onClick={() => { fetchApps(); }}>Fetch all Apps for this Subscription</button>
    </div>
    {progressFetchApps && !devices ? <><br />{'Fetching Apps'}<br /></> : apps ? <><br />Apps<br />{showStatus(apps, showJSON(toggles, apps, toggle, 'apps'), '')}</> : null}

  </div>);
}

export default App;
