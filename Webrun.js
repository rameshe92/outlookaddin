var ruleList, addinMode, LogMode, bNoSigforSub, addSigifNewRec, hideaddininfomsg, multiplesig, domainList, LogModeTemp;
LogMode = false;
LogModeTemp =  false; 
var LogFromAddress = 'demo@sigmails.com'/*'vinit.p@gegadyne.com'*/;
var sPrvFromAddress='';
var sSelectedFromAddress='';
var sNewFromAddress='';
var mutliplesigcontent='';
var sEmbedList = [];
var sEmbedListMultiple = [];
var toEmail = [];
var cctoEmail = [];
var bcctoEmail = [];
var multipleDiv = "mergedsig";
var json_obj = '';
var embedsupportInt = 'false';
var embedsupportExt = 'false';
var InsertSignature = false;
var bExtSigAssigned =  false;
var orgFromAddr = '';
var selectedFromOpt = '0';
var togglemode = '0';
var togglemoderoam;
var removeOutlookSig = 'false';
var bShowAllSig = 'false';
var dropimglinks = 'true';
var bodyType = '';
var hiddenstring = '\u200B\u200B\u200B';
let _mailbox;
let _settings;
var bFirstInsert =-1;
var NewReplyForwardClick =  -1;
var composeType = 'newMail';
var OrgTORecipients = [];
var OrgCCRecipients = [];
var OrgBCCRecipients = [];
var SigInsertMsg = '[Sigsync] Signature has been added.';
var SigInsertMsgEmpty = '[Sigsync] Empty signature added as no templates are assigned!';
var fromDomain = '';
(function () {
	try{
		var item;
		Office.onReady(function() {
			$(document).ready(function () {
				
			});		
		});
		Office.initialize = function() {
		   _mailbox = Office.context.mailbox;
		};
		function onMessageSendHandler(eventArgs) {
			 try {
				eventArgs.completed({ allowEvent: true });
			 } catch (error) {
				console.log('Error in onMessageSendHandler:', error);
			}
			
		}
		function onMessageComposeHandler(event) {	
			if (Office.context) {
				Office.context.roamingSettings.set("bFirstInsert", -1);
				Office.context.roamingSettings.saveAsync(function(result) {
					if (result.status !== Office.AsyncResultStatus.Succeeded) {
					
					} 
				});
			}
			NewReplyForwardClick = 1;
			onMessageComposeHandlerFromTimer(event);
		}		
		function onMessageComposeHandlerFromTimer(event) {		
			const currentDate = new Date();
			if(NewReplyForwardClick !== 1)
				return false;
			try {
				try {
					if (Office.context) {
						
						if (Office.context.mailbox) {
							item = Office.context.mailbox.item;
							if(item) {
								Office.context.mailbox.item.getComposeTypeAsync(function(asyncResult) {
								  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
									composeType = asyncResult.value.composeType;
								  } else {
									console.error(asyncResult.error);
								  }
								});
								
								
								if(item.from) {
									sFromEmailAddress = item.from;
									if(sFromEmailAddress) {
										if(LogModeTemp == true && sFromEmailAddress == LogFromAddress)
											LogRecord("1 OnMSGHnd "+ sFromEmailAddress);

										if (sFromEmailAddress) {
											sFromEmailAddress.getAsync(function (asyncResult) {
												if (asyncResult.status == Office.AsyncResultStatus.Failed) {
													write(asyncResult.error.message);
													if(LogModeTemp == true && sFromEmailAddress == LogFromAddress)
														LogRecord("2 OnMSGHnd "+ asyncResult.error.message);	
												}
												else {
													sNewFromAddress = asyncResult.value.emailAddress;
													if(LogModeTemp == true && sNewFromAddress == LogFromAddress)
														LogRecord("3 OnMSGHnd "+ sNewFromAddress);
												}
											});
										}
									}
								}
							}
						}else{
							if(LogModeTemp == true)
								LogRecord("4 OnMSGHnd Office.context.mailbox is null");
						}
						
					}else{
						if(LogModeTemp == true)
							LogRecord("5 OnMSGHnd Office.context is null");
					}
				} catch(err){
					if(LogModeTemp == true)
						LogRecord("5_1 OnMSGHnd Office.context"+err);
				}
				if(sNewFromAddress == '') {
					if(Office) {
						if(Office.context) {
							if(Office.context.mailbox) {
								if(Office.context.mailbox.userProfile) {
									sNewFromAddress = Office.context.mailbox.userProfile.emailAddress;	
									sSelectedFromAddress = sNewFromAddress;									
								}
							}
						}
					}
				}
				
				if(Office) {
					if(Office.context) {
						if(Office.context.mailbox) {
							if(Office.context.mailbox.userProfile) {
								orgFromAddr = Office.context.mailbox.userProfile.emailAddress;				
							}
						}
					}
				}
				var bServerCall = false;
				if(sSelectedFromAddress != sNewFromAddress){
					sSelectedFromAddress = sNewFromAddress;
					ruleList = false;
					bServerCall = true;
				}
					
				if(orgFromAddr != undefined && orgFromAddr != null)				
					sNewFromAddress = assignFromAddress(selectedFromOpt, sNewFromAddress, orgFromAddr);
				
				if(sPrvFromAddress != sNewFromAddress){
					ruleList=false;
					fromEmail = sNewFromAddress;
					sPrvFromAddress = sNewFromAddress;
					bServerCall = true;
				}
				if(sNewFromAddress) {
					if(sNewFromAddress == LogFromAddress) {
						if(LogModeTemp == true)
							LogRecord("6 OnMSGHnd" + bServerCall);
					}
				} else {
					if(LogModeTemp == true)
							LogRecord("6_1 OnMSGHnd" + currentDate);
				}
				
				
				if(!ruleList && bServerCall == true){
					
					if(fromEmail == LogFromAddress) {
						if(LogModeTemp == true)
							LogRecord("7 OnMSGHnd");
					}
					fromDomain = fromEmail.split('@')[1];
					var xhr = new XMLHttpRequest();
					var indata = "fromEmail=" + fromEmail;
					xhr.open("POST", "https://www.sigsync.com/clientaddin/Insertsig_deploy.php", true);
					xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded"); 			
					xhr.send(indata);
					xhr.onload = function() {
						if(this.responseText) {
							/*var finalStr = LZString.decompressFromBase64(this.responseText);
							var json_obj = JSON.parse(finalStr);*/
							var toRecipients, CCRecipients, BCCRecipients, bExtDomainFoundTmp;
							toRecipients = item.to;
							CCRecipients = item.cc;
							BCCRecipients = item.bcc;
							bExtDomainFoundTmp = false;
							json_obj = JSON.parse(this.responseText);
							
							if (json_obj.log && json_obj.log!=undefined){
								LogMode = json_obj.log;
							}
							if (json_obj.embedsupportInt && json_obj.embedsupportInt == 'true'){
								 embedsupportInt = 'true';
							}
							if (json_obj.embedsupportExt && json_obj.embedsupportExt == 'true'){
								 embedsupportExt = 'true';
							}
							if (json_obj.selectedFromOpt && json_obj.selectedFromOpt!=undefined){
								 selectedFromOpt = json_obj.selectedFromOpt;
							}
							
							if (json_obj.removeOutlookSig && json_obj.removeOutlookSig == 'true'){
								 removeOutlookSig = 'true';
							}
							
							if (json_obj.dropimglinks && json_obj.dropimglinks == 'false'){
								 dropimglinks = 'false';
							}
							toRecipients.getAsync(function (asyncResult) {
								if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								}else {
									for (var i = 0; i < asyncResult.value.length; i++) {
										
										toEmail[i] = asyncResult.value[i].emailAddress;
										
										if(composeType !== 'newMail' && OrgTORecipients.length ==0){
											OrgTORecipients.push(toEmail[i]);
										}
										
										toDomain = toEmail[i].split('@')[1];
										if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
											if(domainList.length > 0) {
												if(domainList.indexOf(toDomain) == -1){
													bExtDomainFoundTmp = true;
												} else {
													if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
														bExtDomainFoundTmp = true;
													}
												}
											} else {
												bExtDomainFoundTmp = true;
											}
										}
									}
								}
							});
							
							CCRecipients.getAsync(function (asyncResult) {
								if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								}else {
									for (var i = 0; i < asyncResult.value.length; i++) {
										
										cctoEmail[i] = asyncResult.value[i].emailAddress;
										if(composeType !== 'newMail' && OrgCCRecipients.length ==0){
											OrgCCRecipients.push(cctoEmail[i]);
										}
										toDomain = cctoEmail[i].split('@')[1];
										if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
											if(domainList.length > 0) {
												if(domainList.indexOf(toDomain) == -1){
													bExtDomainFoundTmp = true;
												} else {
													if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
														bExtDomainFoundTmp = true;
													}
												}
											} else {
												bExtDomainFoundTmp = true;
											}
										}
									}
								}
							});
							
							BCCRecipients.getAsync(function (asyncResult) {
								if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								}else {
									for (var i = 0; i < asyncResult.value.length; i++) {
										
										bcctoEmail[i] = asyncResult.value[i].emailAddress;
										if(composeType !== 'newMail' && OrgBCCRecipients.length ==0){
											OrgBCCRecipients.push(bcctoEmail[i]);
										}
										toDomain = bcctoEmail[i].split('@')[1];
										if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
											if(domainList.length > 0) {
												if(domainList.indexOf(toDomain) == -1){
													bExtDomainFoundTmp = true;
												} else {
													if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
														bExtDomainFoundTmp = true;
													}
												}
											} else {
												bExtDomainFoundTmp = true;
											}
										}
									}
								}
							});
							if(LogMode == true){
								var indata = "From xhrlog=" + fromEmail+ "&Mailbox Support="+ Office.context.diagnostics.version+ "&Platform="+ Office.context.diagnostics.platform +"&responseText WebRun - "+this.responseText;
								LogRecord(indata);
							}
							
							if (json_obj.error){
								ruleList = -1;
								if(json_obj.error == "trial"){
									Office.context.mailbox.item.notificationMessages.addAsync("status", {
											type: "informationalMessage",
											message : "Sigsync signatures trial period has ended. Upgrade your license or contact your administrator",
											icon : "icon32x32",
											persistent: false
										});	
									
								}else if(json_obj.error == "nosig"){
									Office.context.mailbox.item.notificationMessages.addAsync("status", {
										type: "informationalMessage",
										message : "To continue adding signatures, upgrade your Sigsync signatures license or contact your administrator",
										icon : "icon32x32",
										persistent: false
									});	
									
								}else if(json_obj.error == "expired"){
									Office.context.mailbox.item.notificationMessages.addAsync("status", {
										type: "informationalMessage",
										message : "Sigsync signatures subscription has expired. Upgrade your license or contact your administrator",
										icon : "icon32x32",
										persistent: false
									});	
								}else if(json_obj.error == "RuleDataMissing"){
									Office.context.mailbox.item.notificationMessages.addAsync("status", {
										type: "informationalMessage",
										message :"Sender address not selected in any Sigsync signature rule!",
										icon : "icon32x32",
										persistent: false
									});	
								} else {
									Office.context.mailbox.item.notificationMessages.addAsync("status", {
										type: "informationalMessage",
										message : "[Sigsync] Unable to fetch the details. Try again.",
										icon : "icon32x32",
										persistent: false
									});	
								}
							} else {
								if(json_obj.success) {
									if(removeOutlookSig == 'true'){
										if (bodyType === Office.MailboxEnums.BodyType.Text) {
											item.body.setSignatureAsync('',{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
										} else {
											item.body.setSignatureAsync('',{ coercionType: "html" },function(asyncResult) {});
										}
									}
									if(json_obj.addinmode) {
										ruleList = json_obj.success;
										addinMode = json_obj.addinmode;	
										bNoSigforSub = json_obj.bNoSigforSub;	
										hideaddininfomsg = json_obj.hideaddininfomsg;
										domainList = json_obj.dlist;
										if(json_obj.multiplesig)
											multiplesig = json_obj.multiplesig;
										if(json_obj.togglemode){
											togglemode = json_obj.togglemode;
										}
										if(json_obj.addSigifNewRec){
											addSigifNewRec = json_obj.addSigifNewRec;
										}
										
										if (json_obj.ShowAllSig && json_obj.ShowAllSig == 'true'){
											 bShowAllSig = 'true';
										}
										if(ruleList != undefined) {
											var SigIndex = ruleList.length-1;
											if(multiplesig != 1) {
												if(bShowAllSig == 'true') {
													for(t=0;t<ruleList.length;t++){
														value=ruleList[t];
														if(value.ruleapply == 2) {
															SigIndex =  t;
															break;
														}
													}
												}
											}
											
											if(togglemode == '1') {
												_settings = Office.context.roamingSettings;
												togglemoderoam = _settings.get("toggled");
											} else {
												togglemoderoam = 'outlook';
											}
											
											if((togglemoderoam !== 'cloud') && (addinMode != 'noautoinsert') && (addinMode != 'preview')){
												InsertSignature = true;
												if(hideaddininfomsg == 'false') {
													try {
														Office.context.mailbox.item.notificationMessages.addAsync("status", {
															type: "informationalMessage",
															message : "[Sigsync] Inserting signature...",
															icon : "icon32x32",
															persistent: false
														});	
													} catch(err) {
														if(LogMode == true){
															var indata = "Errormsg="+ err.message;
															LogRecord(indata);
														}
													}
												}
												try {
													if(Office.context.diagnostics.platform.toLowerCase() == 'android') {
														ProcessSignatures(event, SigIndex, bExtDomainFoundTmp, item);
													} else {
														Office.context.mailbox.item.body.getTypeAsync(
															function (result) {
																bodyType = result.value;
																if (result.status == Office.AsyncResultStatus.Failed){
																	if(LogMode == true){
																		var indata = "getTypeAsyncresult="+ JSON.stringify(result);
																		LogRecord(indata);
																	}
																} else {
																	ProcessSignatures(event, SigIndex, bExtDomainFoundTmp, item);
																}
															}
														);
													}
													
												} catch(err) {
													if(LogMode == true){
														var indata = "getTypeAsyncresult="+ result.status;
														LogRecord(indata);
													}
												}
											}
										}
									}
								}
							}
						} else {
							Office.context.mailbox.item.notificationMessages.addAsync("status", {
											type: "informationalMessage",
											message : "[Sigsync] Unable to fetch the details. Try again.",
											icon : "icon32x32",
											persistent: false
										});	
						}
					}
				}
				
			} catch(err) {
				LogRecord("Try-Catch. Error " + err);
			}
			
		}
		function ProcessSignatures(event, SigIndex, bExtDomainFoundTmp, item) {
			if(composeType == 'newMail') {
				mutliplesigcontent ='';
				if(multiplesig == 1) { 	 
					for(i=0;i<ruleList.length;i++){
						value=ruleList[i];
						if(value.template!=""){
							mutliplesigcontent += value.template;
							if(value.EmbedDataList.length>0){
								sEmbedListMultiple[i] = value.EmbedDataList;	
							}
						}
						if(value.ruleapply ==2)
							break;
					};
					if(embedsupportExt == 'true') {
						sEmbedList = [].concat(...sEmbedListMultiple);
						setSignaturetoBody(mutliplesigcontent, sEmbedList);
					} else {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(mutliplesigcontent);
						} else {
							item.body.setSignatureAsync(mutliplesigcontent,{ coercionType: "html" },function(asyncResult) {});
						}
					}
					if(mutliplesigcontent == ''){
						SigInsertMsg = SigInsertMsgEmpty;
					}
				} else {
					
					if(embedsupportExt == 'true') {
						setSignaturetoBody(ruleList[SigIndex].template, ruleList[SigIndex].EmbedDataList);
					} else {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(ruleList[SigIndex].template);
						} else {
							item.body.setSignatureAsync(ruleList[SigIndex].template,{ coercionType: "html" },function(asyncResult) {});
						}
					}
					if(ruleList[SigIndex].template == ''){
						SigInsertMsg = SigInsertMsgEmpty;
					}
				}
			} else {
				if(bNoSigforSub == 'true') {
					Office.context.mailbox.item.internetHeaders.setAsync(
						{ "X-Sigsync-Processed": "YesClient" },
						setCallback
					);
					return;
				}
				mutliplesigcontent = '';
				var bSubSigSet = false;
				if(multiplesig == 1){
					for(i=0;i<ruleList.length;i++){
						value=ruleList[i];
						if(value.applyon == 'subemail'){
							if(value.subtemplate!=""){
								bSubSigSet = true;
								mutliplesigcontent += value.subtemplate;
								if(value.SubEmbedDataList.length>0){
									sEmbedListMultiple[i] = value.SubEmbedDataList;	
								}
							}
						} else {
							if(value.template!=""){
									bSubSigSet = true;
								mutliplesigcontent += value.template;
								if(value.EmbedDataList.length>0){
									sEmbedListMultiple[i] = value.EmbedDataList;	
								}
							}
						}
						if(value.ruleapply ==2)
							break;
					};
					if(bSubSigSet == true) {
						if((bExtDomainFoundTmp == true && embedsupportExt == 'true') || (bExtDomainFoundTmp == false && embedsupportInt == 'true')) {
							sEmbedList = [].concat(...sEmbedListMultiple);
							setSignaturetoBody(mutliplesigcontent, sEmbedList);
						} else {
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(mutliplesigcontent);
							} else {
								item.body.setSignatureAsync(mutliplesigcontent,{ coercionType: "html" },function(asyncResult) {});
							}
						}
						if(mutliplesigcontent == ''){
							SigInsertMsg = SigInsertMsgEmpty;
						}
					}
				} else {
					if(ruleList[SigIndex].subtemplatename!='') {
						if((bExtDomainFoundTmp == true && embedsupportExt == 'true') || (bExtDomainFoundTmp == false && embedsupportInt == 'true')) {
							setSignaturetoBody(ruleList[SigIndex].subtemplate, ruleList[SigIndex].SubEmbedDataList);
						} else {
							if (bodyType === Office.MailboxEnums.BodyType.Text) {
								InsertTextSignaturetoBody(ruleList[SigIndex].subtemplate);
							} else {
								item.body.setSignatureAsync(ruleList[SigIndex].subtemplate,{ coercionType: "html" },function(asyncResult) {});
							}
						}
					} else {
						if(ruleList[SigIndex].template!=''){
							if((bExtDomainFoundTmp == true && embedsupportExt == 'true') || (bExtDomainFoundTmp == false && embedsupportInt == 'true')) {
								setSignaturetoBody(ruleList[SigIndex].template, ruleList[SigIndex].EmbedDataList);
							} else {
								if (bodyType === Office.MailboxEnums.BodyType.Text) {
									InsertTextSignaturetoBody(ruleList[SigIndex].template);
								} else {
									item.body.setSignatureAsync(ruleList[SigIndex].template,{ coercionType: "html" },function(asyncResult) {});
								}
								
							}
						}
					}
				}
			}
			OnMessageRecipientsChanged();
			Office.context.mailbox.item.internetHeaders.setAsync(
				{ "X-Sigsync-Processed": "YesClient" },
				setCallback
			);
		}
		function OnMessageRecipientsChanged(event) {
			bExtSigAssigned = false;
			if(NewReplyForwardClick !== 1)
				return false;
			try {
				if(togglemode == '1') {
					_settings = Office.context.roamingSettings;
					togglemoderoam = _settings.get("toggled");
				} else {
					togglemoderoam = 'outlook';
				}
				
				if(togglemoderoam != undefined) {
					if(togglemoderoam == 'cloud') {
						InsertSignature = false;
						Office.context.mailbox.item.internetHeaders.removeAsync(
					  ["X-Sigsync-Processed", "x-sigsync-processed"],
					  function (asyncResult) {  });
					}
				}
							
				if(InsertSignature == true) {
					item = Office.context.mailbox.item;
					var toRecipients, CCRecipients, BCCRecipients, subEmail, bodyEmail;
					toRecipients = item.to;
					CCRecipients = item.cc;
					BCCRecipients = item.bcc;
					subEmail = item.subject;
					bodyEmail = item.body;
					var bExtDomainFound = false;
					try {
						if(toEmail.length >= 1) {
					
							for (var i = 0; i < toEmail.length; i++) {
								toEmail[i] = toEmail[i].toLowerCase();
								toDomain = toEmail[i].split('@')[1];
								if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
									if(domainList.length > 0) {
										if(domainList.indexOf(toDomain) == -1){
											bExtDomainFound = true;
										} else {
											if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
												bExtDomainFoundTmp = true;
											}
										}
									} else {
										bExtDomainFound = true;
									}
								}
							} 
						}
						if(bExtDomainFound == false) {
							bExtSigAssigned = false;
						}
						toEmail.length = 0;
						cctoEmail.length = 0;
						bcctoEmail.length = 0;
					} catch(err) {
					}
					if (toRecipients && bExtSigAssigned == false) {
						
						toRecipients.getAsync(function (asyncResult) {
							if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								
							}else {
								for (var i = 0; i < asyncResult.value.length; i++) {
									count = 1;
									toEmail[i] = asyncResult.value[i].emailAddress;
								}
								
								if(toEmail.length >= 1) {
									if (subEmail){
										subEmail.getAsync(function (asyncResult) {
											if (asyncResult.status == Office.AsyncResultStatus.Failed) {
												
											} else {
												sub = "";
												sub = asyncResult.value;
												if (bodyEmail) {
													bodyEmail.getAsync('text',function (asyncResult) {
														if (asyncResult.status == Office.AsyncResultStatus.Failed) {
															
														} else {
															bodyContent = "";
															bodyContent = asyncResult.value;
														}
														
														setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig);
													});
												}
											}
										});
									}
								}
							}
						});
					} 
					if(CCRecipients && bExtSigAssigned == false) {
						
						CCRecipients.getAsync(function (asyncResult) {
							if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								
							}else {
								
								for (var i = 0; i < asyncResult.value.length; i++) {
									count = 1;
									cctoEmail[i] = asyncResult.value[i].emailAddress;
								}
								
								if(cctoEmail.length >= 1) {
									if (subEmail){
										subEmail.getAsync(function (asyncResult) {
											if (asyncResult.status == Office.AsyncResultStatus.Failed) {
												
											} else {
												sub = "";
												sub = asyncResult.value;
												if (bodyEmail) {
													bodyEmail.getAsync('text',function (asyncResult) {
														if (asyncResult.status == Office.AsyncResultStatus.Failed) {
															
														} else {
															bodyContent = "";
															bodyContent = asyncResult.value;
														}
														toEmail = toEmail.concat(cctoEmail);
														setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig);
													});
												}
											}
										});
									}
								}
							}
						});
					}  
					
					if(BCCRecipients && bExtSigAssigned == false) {
						
						BCCRecipients.getAsync(function (asyncResult) {
							if (asyncResult.status == Office.AsyncResultStatus.Failed) {
								
							}else {
								
								for (var i = 0; i < asyncResult.value.length; i++) {
									count = 1;
									bcctoEmail[i] = asyncResult.value[i].emailAddress;
								}
								
								if(bcctoEmail.length >= 1) {
									if (subEmail){
										subEmail.getAsync(function (asyncResult) {
											if (asyncResult.status == Office.AsyncResultStatus.Failed) {
												
											} else {
												sub = "";
												sub = asyncResult.value;
												if (bodyEmail) {
													bodyEmail.getAsync('text',function (asyncResult) {
														if (asyncResult.status == Office.AsyncResultStatus.Failed) {
															
														} else {
															bodyContent = "";
															bodyContent = asyncResult.value;
														}
														toEmail = toEmail.concat(bcctoEmail);
														setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig);
													});
												}
											}
										});
									}
								}
							}
						});
					}
				}
			} catch(err) {
			  console.log(err);
			}
		}
		
		function setPreview(toEmail, sub, bodyContent, domainList, ruleList, multiplesig) {
			try {
				if (ruleList == "" || toEmail.length <= 0){
					return;
				}
				var bInsertSig =  true;
				if(composeType !== 'newMail' && bNoSigforSub == 'true' && addSigifNewRec == 'true') {
					
					for (var i = 0; i < toEmail.length; i++) {
						var bToFound = false;
						for (var j = 0; j < OrgTORecipients.length; j++) {
							if (toEmail[i].toLowerCase() === OrgTORecipients[j].toLowerCase()) {
								bToFound = true;
								break;
							}
						}
						var bCCFound = false;
						for (var k = 0; k < OrgCCRecipients.length; k++) {
							if (toEmail[i].toLowerCase() === OrgCCRecipients[k].toLowerCase()) {
								bCCFound = true;
								break;
							}
						}
						var bBCCFound = false;
						for (var l = 0; l < OrgBCCRecipients.length; l++) {
							if (toEmail[i].toLowerCase() === OrgBCCRecipients[l].toLowerCase()) {
								bBCCFound = true;
								break;
							}
						}
						
						if (bToFound || bCCFound || bBCCFound) {
							bInsertSig = false;
						} else {
							bInsertSig = true;
							break;
						}
					}
					if(bInsertSig == false) {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody("");
						} else {
							item.body.setSignatureAsync("",{ coercionType: "html" },function(asyncResult) {});
						}
						return;
					}
				} else if(composeType !== 'newMail' && bNoSigforSub == 'true' && addSigifNewRec == 'false') {
					if (bodyType === Office.MailboxEnums.BodyType.Text) {
						InsertTextSignaturetoBody("");
					} else {
						item.body.setSignatureAsync("",{ coercionType: "html" },function(asyncResult) {});
					}
					return;
				}
							
				
				var bSubSigSet = false;
				var toDomain = '';
				var setSig = 0;
				var sigContent = "";
				var firstChar = "";
				var colorCode = "eee";
				var emptyCheck = "false";		
				
				var AddKeywordsList;
				var ExcludeKeywordsList;
				var subSigContent = "";
				var subTemplateName = "";
				var toemailaddress = toEmail[0].toLowerCase();
				var bExtDomainMatched = false;
				sEmbedListMultiple=[];
				for (var i = 0; i < toEmail.length; i++) {
					
					toEmail[i] = toEmail[i].toLowerCase();
					toDomain = toEmail[i].split('@')[1];
					if(toDomain.toLowerCase() !== fromDomain.toLowerCase()) {
						if(domainList.length > 0) {
							if(domainList.indexOf(toDomain) == -1){
								toemailaddress = toEmail[i];
								bExtSigAssigned = true;
								bExtDomainMatched = true;
							} else {
								if(domainList.indexOf(fromDomain.toLowerCase()) == -1){
									toemailaddress = toEmail[i];
									bExtSigAssigned = true;
									bExtDomainMatched = true;
								}
							}
						} else {
							toemailaddress = toEmail[i];
							bExtSigAssigned = true;
							bExtDomainMatched = true;
						}
					}
				} 
								
				if(bExtDomainMatched ==  false){
					bExtSigAssigned = false;
				}
				
				toDomain = toemailaddress.split('@')[1];

				for(j=0;j<ruleList.length;j++){
					value=ruleList[j];
							
					setSig = 0;
					if (value.orgtype != "all") {
						
						if (value.orgtype == "internal" && ((toDomain.toLowerCase() == fromDomain.toLowerCase()) || (domainList.indexOf(toDomain) > -1 && domainList.indexOf(fromDomain) > -1)))
							setSig = 1;
						else if (value.orgtype == "external" && (toDomain.toLowerCase() !== fromDomain.toLowerCase()) && (domainList.length <= 0 || domainList.indexOf(toDomain) == -1 || domainList.indexOf(fromDomain) == -1))
							setSig = 1;
						else {
						
							if(value.addrecipients!=undefined && value.addrecipients!='') {
									
								addlist = JSON.parse(value.addrecipients);									
								
								for(k=0;k<addlist.length;k++){
									addRecipients = JSON.parse(addlist[k]);
									if(addRecipients.rectype == 'listofemails') {
										if(addRecipients.recdata) {
											if(addRecipients.recdata.length > 0) {
											
												for(l=0;l<addRecipients.recdata.length;l++){
													reckeywordarr =addRecipients.recdata[l];
													if(reckeywordarr.recipientemail && reckeywordarr.recipientemail.length > 3) {
														if((toemailaddress == reckeywordarr.recipientemail.toLowerCase()) || (toDomain == reckeywordarr.recipientemail.toLowerCase())){
															setSig = 1;
															break;
														} else if(fnmatch(reckeywordarr.recipientemail, toemailaddress) === true){
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

								for(m=0;m<value.excluderecipients.length;m++){
										
									excludereclist = JSON.parse(value.excluderecipients[m]);
										
									if(excludereclist.rectype == 'listofemails') {
										if(excludereclist.recdata) {
											if(excludereclist.recdata.length > 0) {
												
													for(n=0;n<excludereclist.recdata.length;n++){
														exreckeywordarr = excludereclist.recdata[n];
														
													if(exreckeywordarr.recipientemail && exreckeywordarr.recipientemail.length > 3) {
														if((toemailaddress == exreckeywordarr.recipientemail.toLowerCase()) || (toDomain == exreckeywordarr.recipientemail.toLowerCase())){
															setSig = 0;
															break;
														} else if(fnmatch(reckeywordarr.recipientemail, toemailaddress) === true){
															setSig = 0;
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
					}
			   
					if(setSig == 1) {
						
						if(value.addkeywordlist!=undefined && value.addkeywordlist!='') {
							
							AddKeywordsList = JSON.parse(value.addkeywordlist);
							
							for(o=0;o<AddKeywordsList.length;o++){
								Kvalue = AddKeywordsList[o];
								if(Kvalue.searchtype == "sb" && bodyContent!=undefined && sub!=undefined) {
									if (Kvalue.phrase != "" && ((sub.indexOf(Kvalue.phrase) == -1) || (bodyContent.indexOf(Kvalue.phrase) == -1))) {
										setSig = 0;
										break;
									}
								} else if(Kvalue.searchtype == "s" && sub!=undefined) {
									if (Kvalue.phrase != "" && sub.indexOf(Kvalue.phrase) == -1) {
										setSig = 0;
										break;
									}
								} else if(Kvalue.searchtype == "b" && bodyContent!=undefined) {
									if (Kvalue.phrase != "" && bodyContent.indexOf(Kvalue.phrase) == -1) {
										setSig = 0;
										break;
									}
								}
							}
						}
					}
					if(setSig == 1) {
						if(value.excludekeywordlist!=undefined && value.excludekeywordlist!='') {
						
							ExcludeKeywordsList = JSON.parse(value.excludekeywordlist);
						
								for(p=0;p<ExcludeKeywordsList.length;p++){
									Evalue = ExcludeKeywordsList[p];
								if(Evalue.searchtype == "sb" && bodyContent!=undefined && sub!=undefined) {
									if (Evalue.phrase != "" && ((sub.indexOf(Evalue.phrase) != -1) && (bodyContent.indexOf(Evalue.phrase) != -1))) {
										setSig = 0;
										break;
									}
								} else if(Evalue.searchtype == "s" && sub!=undefined) {
									if (Evalue.phrase != "" && sub.indexOf(Evalue.phrase) != -1) {
										setSig = 0;
										break;
									}
								} else if(Evalue.searchtype == "b" && bodyContent!=undefined) {
									if (Evalue.phrase != "" && bodyContent.indexOf(Evalue.phrase) != -1) {
										setSig = 0;
										break;
									}
								}
							}
						}
					}
             
					if (setSig == 1){					
						
						if(multiplesig != 1) {
							if(composeType == 'newMail'){
							
								if(value.template!="" && value.templatename!= ""){
									sigContent = value.template;
									bSubSigSet = true;
									if(value.EmbedDataList.length>0){
										 sEmbedList = value.EmbedDataList.slice();
									}
								}
							} else {
								if(value.applyon == 'subemail') {
									if(value.subtemplate!="" && value.subtemplatename!= ""){
										sigContent = value.subtemplate;
										bSubSigSet = true;
										if(value.SubEmbedDataList.length>0){
											sEmbedList = value.SubEmbedDataList.slice();
										}
									}
								} else {
									if(value.template!="" && value.templatename!= "") {
										bSubSigSet = true;
										sigContent = sigContent+value.template;
										if(value.EmbedDataList.length>0){
											sEmbedList = value.EmbedDataList.slice();
										}
									}
								}
							}
							
						} else {
							if(composeType == 'newMail') {
								if(value.template!="" && value.templatename!= "") {
									bSubSigSet = true;
									sigContent = sigContent+value.template;
									/*sEmbedList.push(value.EmbedDataList);*/
									sEmbedListMultiple.push(value.EmbedDataList);
								}
							} else {
								if(value.applyon == 'subemail') {
									
									if(value.subtemplate!="" && value.subtemplatename!= ""){
										bSubSigSet = true;
										sigContent = sigContent+value.subtemplate;
										sEmbedListMultiple.push(value.SubEmbedDataList);
									}
								} else {
									if(value.template!="" && value.templatename!= "") {
										bSubSigSet = true;
										sigContent = sigContent+value.template;
										sEmbedListMultiple.push(value.EmbedDataList);
									}
								}
							}
							
						}
						if(value.ruleapply==2) {
							break;
						}
					} else {
						setSig == 0;
						if(value.rulenotapply==2) {
							break;
						}
					}
				}
				if(bSubSigSet == true) {
					if(sEmbedListMultiple.length > 0) {
						sEmbedList = [].concat(...sEmbedListMultiple);
					}
					setSignaturetoBody(sigContent, sEmbedList);
				}
				
			} catch(err) {
			   console.log(err);
			}
		}
		function InsertTextSignaturetoBody(sigContent){
			convertToPlain(sigContent, dropimglinks, function(txtsig){
				var bAppendChar = "";
				var Removepattern = /^[\r\n|\r|\n]/;
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
									
									if (Removepattern.test(bodyText)) {
										bAppendChar = "";
									} else {
										bAppendChar = "\n";
									}	
								}
								modifiedBody = "\n" + hiddenstring + txtsig + hiddenstring + bAppendChar + bodyText;
								item.body.setAsync(modifiedBody, { coercionType: Office.CoercionType.Text }, function (setResult) {
								});
								return;								
							}
						}
						if(bFirstInsert != -1) {
							bodyText = result.value;
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
												if (Removepattern.test(bodyText.substring(sigEndPosition + hiddenstring.length))) {
													bAppendChar = "";
												} else {
													bAppendChar = "\n";
												}	
											}
										}
										modifiedBody = bodyText.substring(0, sigStartPosition) + hiddenstring + txtsig + hiddenstring + bAppendChar + bodyText.substring(sigEndPosition + hiddenstring.length);										
										item.body.setAsync(modifiedBody, { coercionType: Office.CoercionType.Text }, function (setResult) {	});
									}
								}
							}
							if(bInserted == false){
								if(txtsig!='') {
									if (Removepattern.test(txtsig)) {
										bAppendChar = "";
									} else {
										bAppendChar = "\n";
									}	
								}
								item.body.setSignatureAsync(hiddenstring + txtsig + hiddenstring + bAppendChar,{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
							}
						}
					} else {
						if(txtsig!='') {
							if (Removepattern.test(txtsig)) {
								bAppendChar = "";
							} else {
								bAppendChar = "\n";
							}	
						}
						item.body.setSignatureAsync(hiddenstring + txtsig + hiddenstring + bAppendChar,{ coercionType: Office.CoercionType.Text },function(asyncResult) {});
					}
				});
			});
		}
		function setSignaturetoBody(sigContent, EmbedDataListData) {
			try {
				
				if(bExtSigAssigned === true){
					
					if(embedsupportExt == 'false') {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(sigContent);
						} else {
							item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
						}
					} else {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(sigContent);
						} else {	
							if(EmbedDataListData.length>0) {
								ProcessEmbedImagesListRecursive(sigContent, EmbedDataListData, 0,"");
							} else {
								item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
							}
						}
					}
				} else {
					
					if(embedsupportInt == 'false') {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(sigContent);
						} else {
							item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
						}
					} else {
						if (bodyType === Office.MailboxEnums.BodyType.Text) {
							InsertTextSignaturetoBody(sigContent);
						} else {
							if(EmbedDataListData.length>0) {
								 ProcessEmbedImagesListRecursive(sigContent, EmbedDataListData, 0,"");
							} else {
								item.body.setSignatureAsync(sigContent,{ coercionType: "html" },function(asyncResult) {});
							}
						}
					}
				}
			} catch(err) {
				if(LogMode == true){
					var indata = "setSignaturetoBody="+err.message;
					LogRecord(indata);
				}
			}	
		}
		function setCallback(asyncResult){
			try {
				if(hideaddininfomsg == 'false') {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
						type: "informationalMessage",
						message : SigInsertMsg,
						icon : "icon32x32",
						persistent: false
					});
				} else {
					if(LogMode == true){
						var indata = "AsyncResultStatus="+ JSON.stringify(asyncResult.error);
						LogRecord(indata);
					}
				}
				}
			} catch(err) {
				if(LogMode == true){
					var indata = "AsyncResultStatus="+ JSON.stringify(asyncResult.error);
					LogRecord(indata);
				}
			}	
		}
		
		function ProcessEmbedImagesListRecursive(sigContent, file_attachment_arr, index, message) {
			
			
			if (index < file_attachment_arr.length)  {
				
				var file_attachment_obj = file_attachment_arr[index];
				
				if(file_attachment_obj != undefined && file_attachment_obj != null) {
					item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
						if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
							var cleanedBase64Data = file_attachment_obj.imgdata;
							const options = { isInline: true, asyncContext: bodyResult.value };
  
							item.addFileAttachmentFromBase64Async(
								cleanedBase64Data,
								file_attachment_obj.imgname,
								options, (attachResult) => {
								if (attachResult.status === Office.AsyncResultStatus.Succeeded) {	
									if(file_attachment_arr.length == index+1) {
										sigHtml = sigContent.replace(/src/g,'data-url');
										sigHtml = sigHtml.replace(/data-cidpath/g,'src');
										let body = attachResult.asyncContext;
										item.body.setSignatureAsync(sigHtml,{ coercionType: "html" },function(asyncResult) {});
									} else
										ProcessEmbedImagesListRecursive(sigContent, file_attachment_arr, index+1, message);
								}
								});
						} else {
							console.log(bodyResult.error.message);
						}
					});
				} else{
					ProcessEmbedImagesListRecursive(sigContent, file_attachment_arr, index+1, message);
				}
			} 
		}
	
		function assignFromAddress(selectedFromOpt, sFromEmailAddress, orgFromAddr) {
			if(selectedFromOpt != '1'){
				if(orgFromAddr != undefined && orgFromAddr != null)
					return orgFromAddr;
			}
			return sFromEmailAddress;
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
		function setSessionVariable(key, value) {
			if (Office.context) {
				Office.context.roamingSettings.set(key, value);
				Office.context.roamingSettings.saveAsync();
			}
		}

		function getSessionVariable(key) {
			if (Office.context) {
				return Office.context.roamingSettings.get(key);
			} else {
				return true;
			}
		}
		
		function fnmatch(glob, input) {
			var matcher = glob.replace(/\*/g, '.*').replace(/\?/g, '.');
			var regex = new RegExp('^' + matcher + '$');
			return regex.test(input);
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
		
		try {
			Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
			Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
			Office.actions.associate("OnMessageRecipientsChanged", OnMessageRecipientsChanged);
		} catch(err) {
			
			if(LogMode == true){
				var indata = "AsyncResultStatus="+ JSON.stringify(asyncResult.error);
				LogRecord(indata);
			}
		}
		let timerServer = setInterval(onMessageComposeHandlerFromTimer, 3000);		
	} catch(err){
	}
})();
