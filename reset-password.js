/* Today's Basket — password recovery callback (client-side public key only). */
'use strict';
const SB_URL=(window.TB_API_GATEWAY_URL || 'https://kggypoblcagwqqtshixd.supabase.co').replace(/\/+$/, '');
const SB_PUBLISHABLE_KEY='sb_publishable_R19AkJP-laREZyeSf6LkWw_8YTb_UW3';
const $=id=>document.getElementById(id);
const urlHash=new URLSearchParams(location.hash.replace(/^#/,''));
const urlQuery=new URLSearchParams(location.search.replace(/^\?/,''));
const getParam=key=>urlHash.get(key)||urlQuery.get(key);
const errCode=getParam('error_code')||getParam('error');
const errDescription=getParam('error_description')||'';
const proofAtArrival=Boolean(getParam('access_token')||getParam('code')||getParam('token_hash')||getParam('type')==='recovery');
let client=null,verified=false,loading=false;
function status(message,kind='info') { $('resetStatus').textContent=message;$('resetStatus').dataset.kind=kind; }
function cleanURL(){history.replaceState(null,'',location.pathname);}
function showForm(){if(verified)return;verified=true;status('Reset link verified. You can choose a new password.','success');$('resetForm').hidden=false;cleanURL();}
function fail(message){status(message,'error');$('resetForm').hidden=true;}
(async function init(){
 if(errCode){
   const expired=String(errCode).includes('expired')||/expired|invalid|already used/i.test(errDescription);
   fail(expired?'This reset link has expired or already been used. Return to sign in and request another email.':
     `The reset link could not be verified${errDescription?': '+errDescription.replace(/[_+]/g,' '):'. Please request a new link.'}`);
   cleanURL();return;
 }
 if(!proofAtArrival){fail('Please open this page using the reset link sent to your registered email. If you already used the link, request a new one.');return;}
 if(!window.supabase?.createClient){fail('Unable to connect to Supabase. Check your internet connection and reload this link.');return;}
 try{
   client=window.supabase.createClient(SB_URL,SB_PUBLISHABLE_KEY,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:true}});
   client.auth.onAuthStateChange((event,session)=>{
     if(event==='PASSWORD_RECOVERY'&&session)queueMicrotask(showForm);
     if(event==='SIGNED_IN'&&session&&proofAtArrival)queueMicrotask(showForm);
   });
   if(getParam('token_hash')&&getParam('type')==='recovery'){
     const {error}=await client.auth.verifyOtp({token_hash:getParam('token_hash'),type:'recovery'});
     if(error)throw error;
   }
   let {data,error}=await client.auth.getSession();
   if(error||!data.session){
     // PKCE links can contain ?code=. Supabase-js normally exchanges it automatically.
     // If it did not, try exchanging it once here.
     const code=getParam('code');
     if(code){const exchange=await client.auth.exchangeCodeForSession(code);if(exchange.error)throw exchange.error;
       ({data,error}=await client.auth.getSession());}
   }
   if(error||!data?.session)throw Error('This link is invalid, expired, or was already used.');
   showForm();
 }catch(err){fail((err.message||'Unable to verify the reset link.')+' Return to sign in to request another link.');cleanURL();}
})();
$('resetForm').addEventListener('submit',async event=>{
 event.preventDefault();if(!verified||!client||loading)return;
 const password=$('resetNew').value,confirmation=$('resetConfirm').value;
 if(password.length<12||!/[A-Za-z]/.test(password)||!/[0-9]/.test(password)||!/[^A-Za-z0-9]/.test(password)){
   failInline('Please use at least 12 characters with letters, a number and a symbol.');return;}
 if(password!==confirmation){failInline('The passwords do not match. Please check both fields.');return;}
 loading=true;const submit=$('resetSubmit');submit.disabled=true;submit.textContent='Updating password…';
 try{
   const {error}=await client.auth.updateUser({password});if(error)throw error;
   status('Password successfully updated in Supabase. Returning you to the login page…','success');$('resetForm').hidden=true;
   try{await client.auth.signOut();}catch(_){}
   window.setTimeout(()=>{location.replace(new URL('index.html',location.href).href);},1800);
 }catch(err){failInline(err.message||'Password update failed. Please request a new reset link if this continues.');
   loading=false;submit.disabled=false;submit.innerHTML='Update password <span aria-hidden="true">→</span>';}
});
function failInline(message){status(message,'error');$('resetForm').hidden=false;}
