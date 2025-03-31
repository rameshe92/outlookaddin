var sub, bodyContent, domainList, ruleList, newHeight, sigHeight, sigWidth, scaleVal, sigList, addinMode, multiplesig, outlookSup, outlookVer, LogMode, bNoSigforSub, addSigifNewRec;
LogMode = false;
multiplesig = false;
var count = 0;
var sigTagHeight = 0;
var sigPreviewTxtadded = false;
var sPrvFromAddress = '';
var sPrvFromAddressAutoIns = '';
var sNewFromAddress = '';
var sNewFromAddressAutoIns = '';
var toEmail = [];
var bForcedEmpty = false;
var sEmbedList = [];
var sEmbedListMultiple = [];
var divIDList = [];
var orgFromAddr='';
var selectedFromOpt = '0';
var togglemode = '0';
var togglemoderoam;
var cloudmode = '0';
var orgAddinMode = '';
var toggleInsertSig = 0;
var bExtDomainFound = false;
var embedsupportInt = 'false';
var embedsupportExt = 'false';
var removeOutlookSig = 'false';
var bShowAllSig = 'false';
var dropimglinks = 'true';
var bodyType = '';
var hiddenstring = '\u200B\u200B\u200B';
let _mailbox;
let _settings;
var bFirstInsert =-1;
var composeType = 'newMail';
var rangeDisplayVal = '75';
var fromDomain ='';
(function () {
    var item;

    Office.onReady(function() {
		
	});
	Office.initialize = function () {
		_mailbox = Office.context.mailbox;
		_settings = Office.context.roamingSettings;
		Office.context.mailbox.item.getComposeTypeAsync(function(asyncResult) {
		  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			composeType = asyncResult.value.composeType;
		  } else {
			console.error(asyncResult.error);
		  }
		});
		try {
			document.getElementById("signature").style.display = "none";
			document.getElementById("info").style.display = "none";
			document.getElementById("showdiv").style.display = "none";
			document.getElementById("rangeSlider").style.display = "none";
			GetSignatureDetailsFromServer();
		} catch(err) {
			console.log(err.message);
		}
	}
	function OnMessageRecipientsChanged() {
		try {
			if(addinMode == 'preview'){
				try {
					/*setInterval(getAllFields, 1500);*/
					getAllFields();
					if(sigPreviewTxtadded == false) {
						document.getElementById('intro').innerHTML += '<p class="info-msgs info-infoview">Add-in is set to \'<strong><span style="color: #ff0000">Preview Only</span></strong>\' mode.<br /> <a title="Enable / Disable preview mode" href="https://www.sigsync.com/kb/enable-preview-only-mode-for-sigsync-signatures.html" target="_blank" rel="noopener">Click here</a> to change the mode</p>';
						sigPreviewTxtadded = true;
					}
				} catch(err) {
				  
				}
			}
		}catch(err) {
		  
		}
	}
	function GetOutlookVersion(){
		try {						
			outlookVer = Office.context.diagnostics.version;
			var platformcheck = Office.context.diagnostics.platform;
			var buildCheck = outlookVer.split('.');
			var outLookSupport = true;
			/*16.38.614.0 - 16.78.1008.0*/
			if(platformcheck.toLowerCase() == 'mac'){
				if(parseInt(buildCheck[0]) >= 16) {
					if((parseInt(buildCheck[1]) < 59) || (parseInt(buildCheck[2]) < 22031300 && parseInt(buildCheck[1]) == 59)){
						outLookSupport =  false;
					}
				} else {
					outLookSupport =  false;
				}
			} else {
				if(outlookVer != '0.0.0.0'){
					if((parseInt(buildCheck[2]) < 13929 && parseInt(buildCheck[0]) == 16) || (parseInt(buildCheck[0]) < 16)){
						outLookSupport = false;
							
					}
				}
			}
			return outLookSupport;
		} catch(err) {
		  
		}
	}
	
	function ServerCall(fromEmail){
		try{
			fromDomain = fromEmail.split('@')[1];
			var xhr = new XMLHttpRequest();
			var indata = "fromEmail=" + fromEmail;
			xhr.open("POST", "https://www.sigsync.com/clientaddin/Insertsig_deploy.php", true);
			xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 			
			xhr.send(indata);
			xhr.onload = function() {	
				try {
					/*var finalStr = LZString.decompressFromBase64(this.responseText);
					var json_obj = JSON.parse(finalStr);*/
					var json_obj = JSON.parse(this.responseText);
					if (json_obj.log && json_obj.log!=undefined){
						LogMode = json_obj.log;
					}	
					if(LogMode == true){
						var indata = "From xhrlog=" + fromEmail+ "&Mailbox Support="+ Office.context.diagnostics.version+ "&Platform="+ Office.context.diagnostics.platform +"&responseText InsertSig- "+this.responseText;
						LogRecord(indata);
					}
					if (json_obj.error){
						if(json_obj.error == "trial"){
							document.getElementById("info").style.display = "block";
							document.getElementById('info').innerHTML = '<p class="info-msgs info-warning">Sigsync signatures trial period has ended. To continue adding signatures to your emails, upgrade your Sigsync signatures license or contact your administrator.</p>';
						}else if(json_obj.error == "nosig"){
							document.getElementById("info").style.display = "block";
							document.getElementById('info').innerHTML = '<p class="info-msgs info-warning">To continue adding signatures to your emails, upgrade your Sigsync signatures license or contact your administrator.</p>';
						}else if(json_obj.error == "expired"){
							document.getElementById("info").style.display = "block";
							document.getElementById('info').innerHTML = '<p class="info-msgs info-error">Your Sigsync subscription has expired. To continue adding signatures to your emails, upgrade your Sigsync signatures license or contact your administrator.</p>';
						}
					} else {
						if(json_obj.success) {
							if(json_obj.dlist) {								
								ruleList = json_obj.success;
								domainList = json_obj.dlist;
								addinMode = json_obj.addinmode;
								orgAddinMode = json_obj.addinmode;
								disableAddBtn = json_obj.disableaddbtn;
								if(json_obj.multiplesig)
									multiplesig = json_obj.multiplesig;
								
								if(json_obj.selectedFromOpt){
									selectedFromOpt = json_obj.selectedFromOpt;
								}
								if(json_obj.bNoSigforSub){
									bNoSigforSub = json_obj.bNoSigforSub;
								}
								if(json_obj.togglemode){
									togglemode = json_obj.togglemode;
								}
								if(json_obj.cloudmode)
									cloudmode = json_obj.cloudmode;
								if (json_obj.embedsupportInt && json_obj.embedsupportInt == 'true'){
									embedsupportInt = 'true';
								}
								if (json_obj.embedsupportExt && json_obj.embedsupportExt == 'true'){
									 embedsupportExt = 'true';
								}
								if (json_obj.removeOutlookSig && json_obj.removeOutlookSig == 'true'){
									 removeOutlookSig = 'true';
								}
								if (json_obj.ShowAllSig && json_obj.ShowAllSig == 'true'){
									 bShowAllSig = 'true';
								}
								
								if (json_obj.dropimglinks && json_obj.dropimglinks == 'false'){
									 dropimglinks = 'false';
								}
								if(removeOutlookSig == 'true'){
									if (bodyType === Office.MailboxEnums.BodyType.Text) {
										item.body.setSignatureAsync('',{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
									} else {
										item.body.setSignatureAsync('',{ coercionType: "html" },function(asyncResult) {});
									}
								}
								if(addinMode !== 'preview' && togglemode == '1'){
									_settings = Office.context.roamingSettings;
									togglemoderoam = _settings.get("toggled");
									if(_settings.get("toggled") == null || _settings.get("toggled") == undefined) {
										_settings.set("toggled", 'outlook');
										Office.context.roamingSettings.saveAsync(function(result) {
											if (result.status !== Office.AsyncResultStatus.Succeeded) {
												
											} 
										});
									}
									var toggleText = '<p class="modetxt"><input type="radio" name="toggleoption" title="Attach signature while composing email" id="outlook" value="outlook"';
									if(togglemoderoam != 'cloud')
										toggleText += 'checked="checked">';
									else 
										toggleText += '>';
									toggleText += '<label for="outlook" title="Attach signature while composing email">Client Mode</label> <input title="Attach signature on the server" type="radio" name="toggleoption" id="server" value="server"';
									if(togglemoderoam == 'cloud')
										toggleText += 'checked="checked">';
									else 
										toggleText += '>';
									toggleText += '<label for="server" title="Attach signature on the server">Cloud Mode</label>';
									document.getElementById("togglemode").innerHTML = toggleText;
									if(togglemoderoam  == 'cloud'){
										EnableCloudMode();
									}								
								} else if(togglemode == '0'){
									_settings.set("toggled", 'outlook');
									Office.context.roamingSettings.saveAsync(function(result) {
										if (result.status !== Office.AsyncResultStatus.Succeeded) {
											
										}
									});
								}
								if (ruleList == "") {
									 bForcedEmpty = false;
								}
							}
						}
					}
				} catch(err) {
				}				
			}
		} catch(err) {
		}
	}
	
	let timerServer = setInterval(GetSignatureDetailsFromServer, 1000);
	let timerSetsignature = setInterval(waitforItem, 15000);
	
	function GetSignatureDetailsFromServer() {
		try {
			if (Office.context) {		
				if (Office.context.mailbox) {
					item = Office.context.mailbox.item;
					
					sFromEmailAddress = item.from;
					
					if (sFromEmailAddress) {
						sFromEmailAddress.getAsync(function (asyncResult) {
							if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								write(asyncResult.error.message);
							}
							else {
								sNewFromAddress = asyncResult.value.emailAddress;
								orgFromAddr = Office.context.mailbox.userProfile.emailAddress;						
								sNewFromAddress = assignFromAddress(selectedFromOpt, sNewFromAddress, orgFromAddr);
								
								
								sNewFromAddressAutoIns = sNewFromAddress;
								
								if(sPrvFromAddress != sNewFromAddress){
									sPrvFromAddress = sNewFromAddress;
					
									var outLookSupport = GetOutlookVersion();
									if(outLookSupport === false) {
										document.getElementById("info").style.display = "block";
										document.getElementById('info').innerHTML = '<p class="info-msgs info-error">Your Outlook version ('+Office.context.diagnostics.version+') is not compatible with the Sigsync Add-in. If the Signature display is not appearing, then either upgrade your Outlook or contact <a title="Contact Sigsync support" href="https://www.sigsync.com/support.html" target="_blank" rel="noopener">Sigsync support</a></p>';	
									} else {
										ruleList = "";
										domainList = "";
										bForcedEmpty = true;
										ServerCall(sPrvFromAddress);
										timerSetsignature = setInterval(waitforItem, 1000);
									}
								}
								
							}
						});
					}
				}
			}
		} catch(err) {
		}
    }
	function assignFromAddress(selectedFromOpt, sFromEmailAddress, orgFromAddr) {
		if(selectedFromOpt != '1'){
			if(orgFromAddr != undefined && orgFromAddr != null)
				return orgFromAddr;
		}
		return sFromEmailAddress;
	}
	
	function waitforItem() {
		try {
			if (Office.context) {		
				if (Office.context.mailbox) {
					if (Office.context.mailbox.item) {
						clearInterval(timerSetsignature);
						setSignature(domainList, ruleList);
					}
				}
			}
		} catch(err) {
		  if(LogMode == true){
				var indata = "err=" + err.message;
				LogRecord(indata);
			}
		}	
	}
	
    function setSignature(domainList, ruleList, overridepreview) {		
		overridepreview = overridepreview || false;
        if ((ruleList == "") && bForcedEmpty == false){
            try {
				document.getElementById('info').innerHTML += '<p class="info-msgs info-information">Your account is not configured with Sigsync Email Signatures for Office 365. You have to create account with Sigsync and configure your signature.</p><p class="info-msgs info-information">&gt;&gt; <a title="How to add Email signatures" href="https://www.sigsync.com/kb/how-to-add-email-signature.html" target="_blank" rel="noopener">Click here</a> for steps to add signature to your emails.</p><p class="info-msgs info-information">&gt;&gt; <a title="Sigsync signature add-in" href="https://www.sigsync.com/kb/email-signatures-add-in-for-outlook.html" target="_blank" rel="noopener">Click here</a> to know more about Sigsync Outlook Add-in.</p>';
			} catch(err) {
			}
			return;
        }		
		if (ruleList == "") {
			return;
		}
		_settings = Office.context.roamingSettings;
		if(_settings.get("rangeDisplayVal") !== null && _settings.get("rangeDisplayVal") !== undefined) {
			rangeDisplayVal = _settings.get("rangeDisplayVal");
		}
		document.getElementById("rangeSlider").value = rangeDisplayVal;
		if(addinMode == 'preview' && overridepreview == false){
			try {
				/*setInterval(getAllFields, 1500);*/
				getAllFields();
				if(sigPreviewTxtadded == false) {
					document.getElementById('intro').innerHTML += '<p class="info-msgs info-infoview">Add-in is set to \'<strong><span style="color: #ff0000">Preview Only</span></strong>\' mode.<br /> <a title="Enable / Disable preview mode" href="https://www.sigsync.com/kb/enable-preview-only-mode-for-sigsync-signatures.html" target="_blank" rel="noopener">Click here</a> to change the mode</p>';
					sigPreviewTxtadded = true;
				}
			} catch(err) {
			}
		} else{
			try {
				if(sigPreviewTxtadded == false) {
					document.getElementById('intro').innerHTML += '<p class="info-msgs info-infoview"><a title="Enable / Disable preview mode" href="https://www.sigsync.com/kb/enable-preview-only-mode-for-sigsync-signatures.html" target="_blank" rel="noopener">Click here</a> to change the Add-in mode to \'<strong><span style="color: #ff0000">Preview Only</span></strong>\'</p>';
					sigPreviewTxtadded = true;
				}
			
				var sigContent = "";
				var templateName = "";
				var emptyCheck = "true";
				var divName = "";
				var divName2 = "";
				var subSigContent = "";
				var subTemplateName = "";
				var mutliplesigcontent ="";
				var multipleDiv = "mergedsig";
				var datakey="";
				var templatenameList = [];
				divIDList = [];
				document.getElementById('signature').innerHTML = '<div id="emptylist" style="background:#f7f7f7;padding:10px"></div>';
				document.getElementById('emptylist').innerHTML ="";
				document.getElementById("emptylist").style.display = "none";
				document.getElementById("signature").style.display = "none";
				var insertSig = 0;
				
				var sigIndexToInsert = ruleList.length-1;
				
				if(multiplesig != 1) {
					if(bShowAllSig == 'true') {
						for(t=0;t<ruleList.length;t++){
							value=ruleList[t];
							if(value.ruleapply == 2) {
								sigIndexToInsert =  t;
								break;
							}
						}
					}
				}
				ruleList.forEach(function(value, key) {
					if(value.template!="" && value.templatename!= ""){
						sigContent = value.template;
						templateName = value.templatename;
						templateName = templateName.replaceAll(" ","_");
						divName = 'template'+templateName;
						datakey = key;
					}
					if(value.subtemplate!="" && value.subtemplatename!= ""){
						subSigContent = value.subtemplate;
						subTemplateName = value.subtemplatename;
						subTemplateName = subTemplateName.replaceAll(" ","_");
						divName2 = 'subtemplate'+subTemplateName;
						datakey = key;
					}		
					if(multiplesig == 1) {
						if(composeType == 'newMail') {
							if(value.template!="" && value.templatename!= ""){
								mutliplesigcontent += sigContent;
								if(value.EmbedDataList.length>0){
									sEmbedListMultiple[key] = value.EmbedDataList;	
								}
							}
						} else {
							if(value.applyon == 'subemail') {
								if(value.subtemplate!="" && value.subtemplatename!= ""){
									mutliplesigcontent += subSigContent;
									if(value.SubEmbedDataList.length>0){
										sEmbedListMultiple[key] = value.SubEmbedDataList;	
									}
								}
							} else {
								if(value.template!="" && value.templatename!= ""){
									mutliplesigcontent += sigContent;
									if(value.EmbedDataList.length>0){
										sEmbedListMultiple[key] = value.EmbedDataList;	
									}
								}
							}
						}
						if(mutliplesigcontent!='')
							emptyCheck = "false";
					}
					if(multiplesig != 1) {
							if(composeType !== 'newMail') {
							if(value.applyon == 'subemail') {
								if (subSigContent != "") {
							
									if(insertSig == 0 && key == sigIndexToInsert){
										insertSig = 1;
										writeSignature(subSigContent, divName2, datakey, 2, true);
									} else {
										if (!templatenameList.includes(subTemplateName)) {
											templatenameList.push(subTemplateName);
											writeSignature(subSigContent, divName2, datakey, 2, false);
										}
										
									}
									
									subSigContent = "";
									emptyCheck = "false";
									
								} else {
									document.getElementById("signature").style.display = "block";
								}
								if (sigContent != "") {
								
									/*if(insertSig == 0 && key == sigIndexToInsert){
										insertSig = 1;
										writeSignature(sigContent, divName, datakey, 1, true);
									} else {*/
										if (!templatenameList.includes(templateName)) {
											templatenameList.push(templateName);
											writeSignature(sigContent, divName, datakey, 1, false);
										}
										
									/*}*/
									
									sigContent = "";
									emptyCheck = "false";
								} else {
									document.getElementById("signature").style.display = "block";
								}
							} else {
								if (sigContent != "") {
								
									if(insertSig == 0 && key == sigIndexToInsert){
										insertSig = 1;
										writeSignature(sigContent, divName, datakey, 1, true);
									} else {
										if (!templatenameList.includes(templateName)) {
											templatenameList.push(templateName);
											writeSignature(sigContent, divName, datakey, 1, false);
										}
									}
									
									sigContent = "";
									emptyCheck = "false";
								} else {
									document.getElementById("signature").style.display = "block";
								}
							}
						
						} else {
							if (sigContent != "") {
								
								if(insertSig == 0 && key == sigIndexToInsert){
									insertSig = 1;
									writeSignature(sigContent, divName, datakey, 1, true);
								} else {
									if (!templatenameList.includes(templateName)) {
										templatenameList.push(templateName);
										writeSignature(sigContent, divName, datakey, 1, false);
									}
								}
								
								sigContent = "";
								emptyCheck = "false";
							} else {
								document.getElementById("signature").style.display = "block";
							}
						
							if (subSigContent != "") {
								
								if(insertSig == 0 && key == sigIndexToInsert){
									insertSig = 1;
									writeSignature(subSigContent, divName2, datakey, 2, true);
								} else {
									
									if (!templatenameList.includes(subTemplateName)) {
										templatenameList.push(subTemplateName);
										writeSignature(subSigContent, divName2, datakey, 2, false);
									}
								}
								
								subSigContent = "";
								emptyCheck = "false";
								
							} else {
								document.getElementById("signature").style.display = "block";
							}
						}
					} 
					sigContent = "";
					subSigContent = "";
				});
				
				if(multiplesig == 1) {
					if(mutliplesigcontent!='') {
						if(insertSig == 0){
							insertSig = 1;
							writeSignature(mutliplesigcontent, multipleDiv, '', 1, true);
						} else {
							writeSignature(mutliplesigcontent, multipleDiv, '', 1, false);
						}
						
						mutliplesigcontent = '';
					}
				}
				
				if (emptyCheck == "true"){
					
					var emptylist = document.getElementById('emptylist');
					emptylist.style.display = "inline-block";
					emptylist.innerHTML += '<p class="info-msgs info-warning">Signature(s) preview is missing.</p><p class="info-msgs info-warning">1. Ensure that you have created your signature in <a title="Sigsync Email Signatures for Office 365" href="https://www.sigsync.com/kb/how-to-add-email-signature.html" target="_blank" rel="noopener">Sigsync Email Signatures for Office 365</a>.<br />2. Ensure that you have not excluded this sender email address in your <a title="Steps to set Signature Rules" href="https://www.sigsync.com/kb/how-to-set-rules.html" target="_blank" rel="noopener">signature rule</a>.</p>';
				}else{
					document.getElementById("rangeSlider").style.display = "block";
					document.getElementById('emptylist').innerHTML="";
					document.getElementById("emptylist").style.display = "none";
				}
				
			} catch(err) {
			  if(LogMode == true){
					var indata = "setSignatureErr=" + err.message;
					LogRecord(indata);
				}
			}
		}
		var rangeSlider = document.getElementById("rangeSlider");
		var event = new Event('input', {
			bubbles: true,
			cancelable: true,
		});
		rangeSlider.dispatchEvent(event);
    }	
	
	function writeSignature(sigContentTmp, divNameTmp, rulekey, tempType=1, bTriggerclick) {
		try {
			bTriggerclick = bTriggerclick || false;
			sigTagHeight = 0;
			if (!divIDList.includes(divNameTmp)) {
				divIDList.push(divNameTmp);
			
			
				var body = document.getElementsByTagName('body')[0],
				newdiv = document.createElement('div');
				newdiv.id = 'newid';
				body.appendChild(newdiv);
				
				document.getElementById('newid').innerHTML = "<div id="+divNameTmp+" class='table-container'>"+sigContentTmp+"</div>";
						
				var element = document.getElementById(divNameTmp);
				var sigHeight;
				if (element.scrollHeight !== undefined) {
					sigHeight = element.scrollHeight;
				} else {
					sigHeight = 280;
				}
				var sigWidth;
				if (element.scrollWidth !== undefined) {
					sigWidth = element.scrollWidth;
				} else {
					sigWidth = 287;
				}
			
				if(sigWidth <= 0)
					sigWidth = 287;	
				
				if(sigWidth > 287)
					scaleVal = 287 / sigWidth;
				else
					scaleVal = 1;
			
				newHeight = sigHeight*scaleVal;
				removeEl = document.getElementById("newid");
				removeEl.parentNode.removeChild(removeEl);			
				
				var sightml = '<div class="signatureholder"><div id="'+divNameTmp+'"  class="table-container" style="clear:both;"><div class="ms-Grid-row"><div class="ms-Grid-col ms-u-sm12">'+sigContentTmp.replace(/(<br\s*\/?>\s*)+$/, '')+'</div></div></div>';
				if(addinMode != 'preview') {
					if(disableAddBtn == 'true'){
						sightml += '<button class="disabledinsertsig">Add This Signature</button><span class="disabledtxt" style="display:none">You have disabled this option. <a href="https://www.sigsync.com/email-signature/faq.html#disable-add-button" target="_blank">Click here</a> to learn how to enable it.</span>';
						/*<hr style="color:#ddd; width:100%;margin-bottom:5px" />*/
					} else {
						sightml += '<button name="insertbtn" class="insertsig" data-id="'+divNameTmp+'" id="" data-keyval="'+rulekey+'" data-templatetype="'+tempType+'" title="Add this signature template">Add This Signature</button>';
						/*<hr style="color:#ddd; width:100%;margin-bottom:5px" />*/
					}
				}
				sightml += '</div>'; /*<br>*/
				sigTagHeight = sigTagHeight + newHeight;
				document.getElementById('signature').innerHTML += sightml;
				document.getElementById("signature").style.display = "block";	
			}
			if(bTriggerclick == true && (toggleInsertSig == 1 || (addinMode == 'all' && sPrvFromAddressAutoIns != sNewFromAddressAutoIns))){
				if(toggleInsertSig == 1)
					toggleInsertSig = 0;
				sPrvFromAddressAutoIns = sNewFromAddressAutoIns;
				if(composeType == 'newMail' || (composeType !== 'newMail' && bNoSigforSub !== 'true')) {
					var elements = document.querySelectorAll('.insertsig[data-id="' + divNameTmp + '"]');
					if (elements.length > 0) {
						elements[0].click();
					}
				} else {
					Office.context.mailbox.item.internetHeaders.setAsync(
						{ "X-Sigsync-Processed": "YesClient" }					
					);
				}
			}
			
			
		} catch(err) {
		  if(LogMode == true){
			var indata = "writeSignature=" + err.message;
			LogRecord(indata);
			}
		}
	}
	
    function write(message) {        
		console.log(message);
    }	
	function EnableCloudMode(){
		if (bodyType === Office.MailboxEnums.BodyType.Text) {
			item.body.setSignatureAsync('',{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
		} else {
			item.body.setSignatureAsync('',{ coercionType: "html" },function(asyncResult) {});
		}
		addinMode = 'preview';
		Office.context.mailbox.item.internetHeaders.removeAsync(
		  ["X-Sigsync-Processed", 'x-sigsync-processed'],
		  function (asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			  
			} 
		  }
		);
	}		
	function getAllFields() {
        try {
			if(Office.context.mailbox != undefined) {
				item = Office.context.mailbox.item;
				var toRecipients, subEmail, bodyEmail;
				toRecipients = item.to;
				subEmail = item.subject;
				bodyEmail = item.body;
				
				if (toRecipients)
					toRecipients.getAsync(function (asyncResult) {
						if (asyncResult.status == Office.AsyncResultStatus.Failed) {
							/*write(asyncResult.error.message);*/
						}
						else {
							toEmail.length = 0;
							for (var i = 0; i < asyncResult.value.length; i++) {
								count = 1;
								toEmail[i] = asyncResult.value[i].emailAddress;
							}
						}
					});
					
				/*if (count == 0) {
					document.getElementById('info').innerHTML = '<p class="info-msgs info-information">Enter an \"Email address\" in <b>\"To\"</b> field to preview the signature.</p>';
					document.getElementById("info").style.display = "block";
					document.getElementById("signature").style.display = "none";
				} else {
					document.getElementById('info').innerHTML = "";
					document.getElementById("info").style.display = "none";
					count = 0;
				}*/
	
				if (subEmail)
					subEmail.getAsync(function (asyncResult) {
						if (asyncResult.status == Office.AsyncResultStatus.Failed) {
							write(asyncResult.error.message);
						} else {
							sub = "";
							sub = asyncResult.value;
						}
					});
			
				if (bodyEmail)
					bodyEmail.getAsync('text',function (asyncResult) {
						if (asyncResult.status == Office.AsyncResultStatus.Failed) {
							write(asyncResult.error.message);
						} else {
							bodyContent = "";
							bodyContent = asyncResult.value;
						}
					});
	
				setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig);
			}
		} catch(err) {
		   if(LogMode == true){
			var indata = "getAllFields=" + err.message;
			LogRecord(indata);
			}
		}
    }
	
	function setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig) {
        try {
			if (ruleList == ""){
				document.getElementById('info').innerHTML += '<p class="info-msgs info-information">Your account is not configured with Sigsync Email Signatures for Office 365. You have to create account with Sigsync and configure your signature.</p><p class="info-msgs info-information">&gt;&gt; <a title="How to add Email signatures" href="https://www.sigsync.com/kb/how-to-add-email-signature.html" target="_blank" rel="noopener">Click here</a> for steps to add signature to your emails.</p><p class="info-msgs info-information">&gt;&gt; <a title="Sigsync signature add-in" href="https://www.sigsync.com/kb/email-signatures-add-in-for-outlook.html" target="_blank" rel="noopener">Click here</a> to know more about Sigsync Outlook Add-in.</p>';
				return;
			}
			document.getElementById("rangeSlider").style.display = "block";
			var toDomain = '';
			var setSig = 0;
			var sigContent = "";
			var firstChar = "";
			var colorCode = "eee";
			var emptyCheck = "false";		
			
			var AddKeywordsList;
			var ExcludeKeywordsList;
			var divName2 = "";
			var subSigContent = "";
			var subTemplateName = "";
			sigTagHeight = 0;
			sigList = "";
			
			document.getElementById('signature').innerHTML = '<div id="emptylist" style="background:#f7f7f7;padding:10px"></div>';
			
			document.getElementById('emptylist').innerHTML ="";
			document.getElementById("emptylist").style.display = "none";
			document.getElementById("signature").style.display = "none";
			if(toEmail.length <=0) {
				setSignature(domainList, ruleList, true);
			} else {
				for (var i = 0; i < toEmail.length; i++) {
					toEmail[i] = toEmail[i].toLowerCase();
					toDomain = toEmail[i].split('@')[1];
				
					firstChar = toEmail[i].charAt(0);
					colorCode = firstChar.charCodeAt(0)%26;
					colorCode = getColorCode(colorCode);
					var divName = "";
					var templateName = "";	
					
					ruleList.forEach(function(value, key) {
						setSig = 0;
						if (value.orgtype != "all") {
							if (value.orgtype == "internal" && ((toDomain.toLowerCase() == fromDomain.toLowerCase()) || (domainList.indexOf(toDomain) > -1 && domainList.indexOf(fromDomain) > -1)))
								setSig = 1;
							else if (value.orgtype == "external" && (toDomain.toLowerCase() !== fromDomain.toLowerCase()) && (domainList.length <= 0 || domainList.indexOf(toDomain) == -1 || domainList.indexOf(fromDomain) == -1))
								setSig = 1;
							else {
							
								if(value.addrecipients!=undefined && value.addrecipients!='') {
									var addlist = JSON.parse(value.addrecipients);
									for (var reckey in addlist) {
										if (addlist.hasOwnProperty(reckey)) {
											var recval = addlist[reckey];
											var addRecipients = JSON.parse(recval);
											if (addRecipients.rectype === 'listofemails') {
												if (addRecipients.recdata && addRecipients.recdata.length > 0) {
													for (var i = 0; i < addRecipients.recdata.length; i++) {
														var reckeywordarr = addRecipients.recdata[i];

														if (reckeywordarr.recipientemail && reckeywordarr.recipientemail.length > 3) {
															if ((toEmail[i] === reckeywordarr.recipientemail.toLowerCase()) || (toDomain === reckeywordarr.recipientemail.toLowerCase())) {
																setSig = 1;
																break;
															} else if (fnmatch(reckeywordarr.recipientemail, toEmail[i]) === true) {
																setSig = 1;
																break;
															}
														}
													}
												}
											}
										}
									}
								}
							}
						} else
							setSig = 1;
						if(setSig == 1) {
							if(value.donotaddruletype) {
							   if (value.donotaddruletype == "internal" && ((toDomain.toLowerCase() == fromDomain.toLowerCase()) || domainList.indexOf(toDomain) > -1)) {
									setSig = 0;
							   }else if (value.donotaddruletype == "external" && (toDomain.toLowerCase() !== fromDomain.toLowerCase()) && (domainList.length <= 0 || domainList.indexOf(toDomain) == -1)) {
								   setSig = 0;
							   } 
							}	
							if(setSig == 1) {
								if(value.excluderecipients!=undefined && value.excluderecipients!='') {
									value.excluderecipients.forEach(function(exrecval, exreckey) {
										var excludereclist = JSON.parse(exrecval);
										if (excludereclist.rectype === 'listofemails') {
											if (excludereclist.recdata && excludereclist.recdata.length > 0) {
												excludereclist.recdata.forEach(function(exreckeywordarr, exrecdatakey) {
													if (exreckeywordarr.recipientemail && exreckeywordarr.recipientemail.length > 3) {
														if ((toEmail[i] === exreckeywordarr.recipientemail.toLowerCase()) || (toDomain === exreckeywordarr.recipientemail.toLowerCase())) {
															setSig = 0;
															return;
														} else if (fnmatch(reckeywordarr.recipientemail, toEmail[i]) === true) {
															setSig = 0;
															return;
														}
													}
												});
											}
										}
									});

								}
							}
						}
				   
						if(setSig == 1) {
							if(value.addkeywordlist!=undefined && value.addkeywordlist!='') {
								var AddKeywordsList = JSON.parse(value.addkeywordlist);
								AddKeywordsList.forEach(function(Kvalue, Kkey) {
									if (Kvalue.searchtype === "sb" && bodyContent !== undefined && sub !== undefined) {
										if (Kvalue.phrase !== "" && (sub.indexOf(Kvalue.phrase) === -1 || bodyContent.indexOf(Kvalue.phrase) === -1)) {
											setSig = 0;
											return;
										}
									} else if (Kvalue.searchtype === "s" && sub !== undefined) {
										if (Kvalue.phrase !== "" && sub.indexOf(Kvalue.phrase) === -1) {
											setSig = 0;
											return;
										}
									} else if (Kvalue.searchtype === "b" && bodyContent !== undefined) {
										if (Kvalue.phrase !== "" && bodyContent.indexOf(Kvalue.phrase) === -1) {
											setSig = 0;
											return;
										}
									}
								});
							}
						}
						if(setSig == 1) {
							if(value.excludekeywordlist!=undefined && value.excludekeywordlist!='') {
								var ExcludeKeywordsList = JSON.parse(value.excludekeywordlist);
								ExcludeKeywordsList.forEach(function(Evalue, Ekey) {
									if (Evalue.searchtype === "sb" && bodyContent !== undefined && sub !== undefined) {
										if (Evalue.phrase !== "" && (sub.indexOf(Evalue.phrase) !== -1 && bodyContent.indexOf(Evalue.phrase) !== -1)) {
											setSig = 0;
											return;
										}
									} else if (Evalue.searchtype === "s" && sub !== undefined) {
										if (Evalue.phrase !== "" && sub.indexOf(Evalue.phrase) !== -1) {
											setSig = 0;
											return;
										}
									} else if (Evalue.searchtype === "b" && bodyContent !== undefined) {
										if (Evalue.phrase !== "" && bodyContent.indexOf(Evalue.phrase) !== -1) {
											setSig = 0;
											return;
										}
									}
								});

							}
						}
				 
						if (setSig == 1){					
							
							if(multiplesig != 1) {
								if(value.template!="" && value.templatename!= ""){
									sigContent = value.template;
									templateName = value.templatename;
									divName = 'template'+key;
								}
							} else {
								if(composeType == 'newMail') {
									if(value.template!="" && value.templatename!= "") {
										
										sigContent += value.template;
										divName += 'mtemplate'+key;
										templateName += value.templatename;
									}
								} else {
									if(value.applyon == 'subemail') {
										if(value.subtemplate!="" && value.subtemplatename!= ""){
											sigContent += value.subtemplate;
											divName += 'mtemplate'+key;
											templateName += value.subtemplatename;
										}
									} else {
										if(value.template!="" && value.templatename!= "") {
											sigContent += value.template;
											divName += 'mtemplate'+key;
											templateName += value.templatename;
										}
									}
								}
							}
							if(value.ruleapply==2) {
								return false;
							}
						} else {
							setSig == 0;
							if(value.rulenotapply==2) {
								return false;
							}
						}
					});
				
					if (sigContent != "") {
						
						if(sigList.indexOf(templateName+';') == -1){
							sigList = sigList+templateName+';';
							var body = document.getElementsByTagName('body')[0],
							newdiv = document.createElement('div');
							newdiv.id = 'newid';
							body.appendChild(newdiv);
							document.getElementById('newid').innerHTML = "<div id="+divName+" class=table-container>"+sigContent+"</div>";
						
							var element = document.getElementById(divName);
							var sigHeight;
							if (element.scrollHeight !== undefined) {
								sigHeight = element.scrollHeight;
							} else {
								sigHeight = 280;
							}
							var sigWidth;
							if (element.scrollWidth !== undefined) {
								sigWidth = element.scrollWidth;
							} else {
								sigWidth = 287;
							}

							if(sigWidth > 287)
								scaleVal = 287 / sigWidth;
							else
								scaleVal = 1;
							newHeight = sigHeight*scaleVal;
							removeEl = document.getElementById("newid");
							removeEl.parentNode.removeChild(removeEl);
							document.getElementById('signature').innerHTML += '<div id="'+divName+'list"><p class="icon-holder"><span class="first-icon" style="background:#'+colorCode+'" title="'+toEmail[i]+'">'+firstChar+'</span></p></div>';
							document.getElementById('signature').innerHTML += '<div id="'+divName+'" style="transform:scale('+scaleVal+');transform-origin:left top;height: '+newHeight+'px;">'+sigContent.replace(/(<br\s*\/?>\s*)+$/, '')+'</div>';
							sigTagHeight = sigTagHeight + newHeight;
							document.getElementById("signature").setAttribute('style','height:'+sigTagHeight+'px;');
							document.getElementById("signature").style.display = "block";
						}
						else {
							var emptylist = document.getElementById('emptylist');
							var htmlString = '<p class="icon-holder"><span class="first-icon" style="background:#'+colorCode+'" title="'+toEmail[i]+'">'+firstChar+'</span></p>';
							emptylist.innerHTML += htmlString;
						}

						sigContent = "";
					} else {
						document.getElementById("signature").style.display = "block";
						var emptylist = document.getElementById('emptylist');
						var htmlString = '<p class="icon-holder"><span class="empty-icon" style="background:#'+colorCode+';display:inline;padding:2px 8px 5px" title="'+toEmail[i]+'">'+firstChar+'</span>'+toEmail[i]+'</p>';
						emptylist.innerHTML += htmlString;
						emptyCheck = "true";
					}
				}
			}
			if (emptyCheck == "true"){
				document.getElementById("emptylist").style.display = "inline-block";
				var emptylist = document.getElementById('emptylist');
				var htmlString = '<p class="info-msgs info-warning">Signature(s) preview is missing.</p><p class="info-msgs info-warning">1. Ensure that you have created your signature in <a title="Sigsync Email Signatures for Office 365" href="https://www.sigsync.com/kb/how-to-add-email-signature.html" target="_blank" rel="noopener">Sigsync Email Signatures for Office 365</a>.<br />2. Ensure that you have not excluded this sender email address in your <a title="Steps to set Signature Rules" href="https://www.sigsync.com/kb/how-to-set-rules.html" target="_blank" rel="noopener">signature rule</a>.</p>';
				emptylist.innerHTML += htmlString;
			}
			else{
				document.getElementById("emptylist").innerHTML="";
				document.getElementById("emptylist").style.display = "none";
			}	
		} catch(err) {
		   if(LogMode == true){
			var indata = "PreviewArr=" + err.message;
			LogRecord(indata);
			}
		}
    }
	
	function getColorCode(index) {
		var COLORCODE = {0:'7551d3',1:'7550d1',2:'2c85ec',3:'0bb3b2',4:'2080bd',5:'da5c38',6:'8a4183',7:'3c8deb',8:'8C8775',9:'4ea9ce', 10:'267494', 11:'b4a11e',12:'bc4726',13:'8b4884',14:'5c3daa',15:'7b4676',16:'6d8686',17:'ad6f3e',18:'a38641',19:'963171',20:'4096b9',21:'809827', 22:'b30414', 23:'83a539',24:'a73644',25:'297fa7',26:'b19935'};
		return COLORCODE[index];
	}
	document.addEventListener("click", function(event) {
		if (event.target && event.target.matches('input[name="toggleoption"]')) {
			document.getElementById('errorid').style.display = "none";
			if (event.target.value === 'outlook') {
				addinMode = orgAddinMode;
				document.getElementById("info").style.display="none";
				toggleInsertSig = 1;
				_settings = Office.context.roamingSettings;
				_settings.set("toggled", 'outlook');
				Office.context.roamingSettings.saveAsync(function(result) {
				  if (result.status !== Office.AsyncResultStatus.Succeeded) {
				  }
				});
				document.getElementById("modelbl").innerHTML = "'Client Mode'";
				document.getElementById('togglealert').style.display = "block";
				setTimeout(function() {hideErrorMessage('togglealert');}, 10000);
			} else {
				if (event.target.value === 'server') {
					if(cloudmode == '1'){
						_settings = Office.context.roamingSettings;
						 _settings.set("toggled", 'cloud');
						Office.context.roamingSettings.saveAsync(function(result) {
						  if (result.status !== Office.AsyncResultStatus.Succeeded) {
							
						  }
						});
						document.getElementById("modelbl").innerHTML = "'Cloud Mode'";
						document.getElementById('togglealert').style.display = "block";
						EnableCloudMode();
						setTimeout(function() {hideErrorMessage('togglealert');}, 10000);
					} else {
						document.getElementById('errorid').style.display = "block";
						document.getElementById('togglealert').style.display = "none";
						document.getElementById('server').checked = false;
						document.getElementById('outlook').checked = true;
						setTimeout(function() {hideErrorMessage('errorid');}, 10000);
					}
				}
			}
			setSignature(domainList, ruleList);
		}
	});
	function hideErrorMessage(div_id) {
		var errorMessageDiv = document.getElementById(div_id);
		errorMessageDiv.style.display = 'none';
	}
	function convertToPlain(html, dropimglinks, callback){
		try {
			var indata = "htmltext="+ encodeURIComponent(html) + "&dropimglinks=" + dropimglinks;
			var xhrlog1 = new XMLHttpRequest();
			xhrlog1.open("POST", "https://www.sigsync.com/clientaddin/converthtmlforAddin.php", true);
			xhrlog1.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 			
			xhrlog1.send(indata);
			xhrlog1.onload = function() {
				if(this.responseText) {
					if (this.status === 200) {
						callback(this.responseText); 
					}
				}
			}
		} catch(err) {
		  
		}
	}
	document.addEventListener("input", function(event) {
		if(event.target.matches('#rangeSlider')) {
			var slider = document.getElementById("rangeSlider");
			rangeDisplayVal = slider.value;
			var value = rangeDisplayVal;
			Office.context.roamingSettings.set("rangeDisplayVal", rangeDisplayVal);
			Office.context.roamingSettings.saveAsync(function(result) {
				if (result.status !== Office.AsyncResultStatus.Succeeded) {
					
				} 
			});
			var scaleValue = value / 100;
			var signatureHolder = document.querySelector("#signature");
			signatureHolder.style.transform = "scale(" + scaleValue + ")";
			
			var btnContainers = document.querySelectorAll(".insertsig");
			btnContainers.forEach(function(btnContainer) {
				if(scaleValue < 0.6) {
					scaleval  = (2-scaleValue);
					btnContainer.style.transform = "scale(" + (scaleval) + ")";
					btnContainer.style.transformOrigin = "0px 0px";
				} else {
					btnContainer.style.transform = "inherit";
				}
			});
		}
	});
	document.addEventListener("click", function(event) {
		if (event.target.matches('.disabledinsertsig')) {
			var disabledElements = document.querySelectorAll(".disabledtxt");
			disabledElements.forEach(function(element) {
				element.style.display = "block";
			});
		}
	});
	document.addEventListener("click", function(event) {
		if (event.target.matches('.insertsig')) {
			
			var dataId = event.target.getAttribute('data-id');
			var datakey = event.target.getAttribute('data-keyval');
			var dataTemplateType = event.target.getAttribute('data-templatetype');
			var sigContent = document.getElementById(dataId).innerHTML;
			var Embedlist;
			item.body.getTypeAsync(
			function (result) {
				if (result.status == Office.AsyncResultStatus.Failed){
					writebody(result.error.message);
				}else {
					bodyType = result.value;
					if(datakey!='') {
						if(dataTemplateType ==2)
							Embedlist = ruleList[datakey].SubEmbedDataList;
						else
							Embedlist = ruleList[datakey].EmbedDataList;
					} else {
						if(sEmbedListMultiple.length >0){
							Embedlist = [].concat(...sEmbedListMultiple);
						}
					}
				
					for (var i = 0; i < toEmail.length; i++) {
						toEmail[i] = toEmail[i].toLowerCase();
						toDomain = toEmail[i].split('@')[1];
						if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
							if(domainList.length <= 0 || domainList.indexOf(toDomain) == -1){
								bExtDomainFound = true;
							} else {
								if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
									bExtDomainFound = true;
								}
							}
						}
					}
					if(bExtDomainFound == true) {
						if(embedsupportExt == 'true'){
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(sigContent);
							} else {
								if(Embedlist!=undefined && Embedlist.length>0) {
									ProcessEmbedImagesListRecursive(sigContent, Embedlist, 0,"");
								} else {
									item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
								}
							}
						} else {
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(sigContent);
							} else {
								item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
							}
						}
						
					} else {
						
						if(embedsupportInt == 'true'){
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(sigContent);
							} else {
								if(Embedlist!=undefined && Embedlist.length>0) {
									ProcessEmbedImagesListRecursive(sigContent, Embedlist, 0,"");
								} else {
									item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
								}
							}
						} else {
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(sigContent);						
							} else {
								item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
							}
						}
					}
					Office.context.mailbox.item.internetHeaders.setAsync(
						{ "X-Sigsync-Processed": "YesClient" }					
					);
				}
			});
		}
	});
	function fnmatch(glob, input) {
	  var matcher = glob.replace(/\*/g, '.*').replace(/\?/g, '.');
	  var regex = new RegExp('^' + matcher + '$');
	  return regex.test(input);
	}
	function InsertTextSignaturetoBody(sigContent){
		convertToPlain(sigContent, dropimglinks, function(txtsig){
			var bAppendChar = "";
			item.body.getAsync(Office.CoercionType.Text, function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					bodyText  = result.value;
					bFirstInsert = Office.context.roamingSettings.get("bFirstInsert");
					if(bFirstInsert == -1){
						bFirstInsert = 0;
						Office.context.roamingSettings.set("bFirstInsert", bFirstInsert);
						Office.context.roamingSettings.saveAsync(function(result) {
							if (result.status !== Office.AsyncResultStatus.Succeeded) {
								
							} 
						});
						if(bodyText!='') { 
							bFirstInsert = 1;
							Office.context.roamingSettings.set("bFirstInsert", bFirstInsert);
							Office.context.roamingSettings.saveAsync(function(result) {
								if (result.status !== Office.AsyncResultStatus.Succeeded) {
									
								} 
							});
							bodyText = bodyText.replace(/\u200B\u200B\u200B/g, "");
							if(bodyText!='') {
								if (/^[\r\n|\r|\n]/.test(bodyText)) {
									bAppendChar = "";
								} else {
									bAppendChar = "\n";
								}	
							}
							modifiedBody = hiddenstring + txtsig + hiddenstring + bAppendChar + bodyText;
							item.body.setAsync(modifiedBody, { coercionType: Office.CoercionType.Text }, function (setResult) {
							});
							return;								
						}
					}
					if(bFirstInsert != -1) {
						var sigStartPosition=-1;
						var sigEndPosition=-1;
						var bInserted = false;
						if(bodyText!=''){
							sigStartPosition = bodyText.indexOf(hiddenstring);
							if (sigStartPosition !== -1 && bodyText.length > (sigStartPosition + hiddenstring.length)) {
								sigEndPosition = bodyText.indexOf(hiddenstring, sigStartPosition + hiddenstring.length);
							}
							
							if (sigStartPosition !== -1 && sigEndPosition !== -1) {
								if (sigStartPosition !== sigEndPosition) {
									bInserted = true;
									if (bodyText.length > (sigEndPosition + hiddenstring.length)) {
										if(bodyText.substring(sigEndPosition + hiddenstring.length)!='') {
											if (/^[\r\n|\r|\n]/.test(bodyText.substring(sigEndPosition + hiddenstring.length))) {
												bAppendChar = "";
											} else {
												bAppendChar = "\n";
											}	
										}
									}
									var bodyTextstart = bodyText.substring(0, sigStartPosition).replace(/\u200B\u200B\u200B/g, "");
									var bodyTextEnd = bodyText.substring(sigEndPosition + hiddenstring.length).replace(/\u200B\u200B\u200B/g, "");
									modifiedBody = bodyTextstart + hiddenstring + txtsig + hiddenstring + bAppendChar + bodyTextEnd;
									item.body.setAsync(modifiedBody, { coercionType: Office.CoercionType.Text }, function (setResult) {	});
									return;
								}
							}
						}
						if(bInserted == false){
							if(txtsig!='') {
								if (/[\r\n|\r|\n]$/.test(txtsig)) {
									bAppendChar = "";
								} else {
									bAppendChar = "\n";
								}	
							}
							bodyText = bodyText.replace(/\u200B\u200B\u200B/g, "");
							item.body.setAsync(bodyText, { coercionType: Office.CoercionType.Text }, function (setResult) {	
								item.body.setSignatureAsync(hiddenstring + txtsig + hiddenstring + bAppendChar,{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
							});
						}
					}
				} else {
					if(txtsig!='') {
						if (/[\r\n|\r|\n]$/.test(txtsig)) {
							bAppendChar = "";
						} else {
							bAppendChar = "\n";
						}	
					}
					bodyText = bodyText.replace(/\u200B\u200B\u200B/g, "");
					item.body.setAsync(bodyText, { coercionType: Office.CoercionType.Text }, function (setResult) {	
						item.body.setSignatureAsync(hiddenstring + txtsig + hiddenstring + bAppendChar,{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
					});
				}
			});
		});
	}
	
	
	document.addEventListener("click", function(event) {
		if (event.target.matches('.info-txt')) {
			 event.target.classList.toggle('active');
			var x = document.getElementById("showdiv");
			var y = document.getElementById("infotxt");
			if (x.style.display === "none") {
				y.innerHTML = "Click here to hide the steps";
				x.style.display = "block";
				y.classList.add("active");
			} else {
				y.innerHTML = "Click here to know Add-in configuration steps";
				x.style.display = "none";
				y.classList.remove("active");
			}
		}
	});
	function ProcessEmbedImagesListRecursive(sigContent, file_attachment_arr, index, message) {
		const options = { isInline: true }; 
		if (index < file_attachment_arr.length)  {
			var file_attachment_obj = file_attachment_arr[index];
			item.addFileAttachmentFromBase64Async(
				file_attachment_obj.imgdata,
				file_attachment_obj.imgname,
				options,
				function(result1){						
					if(file_attachment_arr.length == index+1) {
						sigHtml = sigContent.replace(/src/g,'data-url');
						sigHtml = sigHtml.replace(/data-cidpath/g,'src');
						item.body.setSignatureAsync(sigHtml,{ coercionType: "html" },function(asyncResult) {});
					} else
						ProcessEmbedImagesListRecursive(sigContent, file_attachment_arr, index+1, message);
				}
			);
		} 
	}
	function LogRecord(indata) {
		try {
			var xhrlog = new XMLHttpRequest();
			xhrlog.open("POST", "https://www.sigsync.com/clientaddin/writelog_deploy.php", true);
			xhrlog.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 			
			xhrlog.send(indata);
		} catch(err) {
		}
	}
	
})();