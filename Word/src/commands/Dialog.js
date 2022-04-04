

  var scopeindex=null;
  var vjson="";
  var classscore="";
  var sensescore="";
  var  cscoremustbe=null
  var sscoremustbe=null;
  var newclassguid=null;
  var newsenseguid=null;
    Office.onReady()
     .then (function()
        {

            Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,onMessageFromParent);
            //console.log("dialog ready");
            msgtoparent("Ready",null);
           // var readymessage= JSON.stringify("Ready");
          //  Office.context.ui.messageParent(readymessage);
           $('#classList').change(function(e)
           {
             //console.log("selected class guid :"+($(this).val()));
             newclassguid = ($(this).val());
             newsenseguid = $('#sensitivityList').val();
             //console.log("sense selected val: "+ $('#sensitivityList').val());
             //console.log("new class guid :" + newclassguid);
             updateheader();
             
           });
           $('#clsButton').click(function(e)
           {
             //console.log("close clicked");
             msgtoparent('close','closeme');
           })
           $('#sensitivityList').change(function(e)
           {
            newsenseguid = ($(this).val());
            newclassguid = $('#classList').val();
            updateheader();
             
           });
        }
     );

     function msglogger(msg)
     {
       let txt = document.getElementById("logger").value;
       document.getElementById("logger").value = txt+"--"+msg;
     }

 function onMessageFromParent(arg)
     {
         var messagefromparent = JSON.parse(arg.message);
       //  msglogger(arg.message);
         if(messagefromparent.messageType =="vket Json")
          {
           //console.log("message from parent :"+messagefromparent.data["Policies"].length);
           vjson = messagefromparent.data;
          } else if (messagefromparent.messageType=="Politika")
          {
              //console.log("Politika :"+ messagefromparent.data);
              document.getElementById("politika").value=messagefromparent.data;
            //  $('#prg').text(messagefromparent.data);
          } else if (messagefromparent.messageType =="Kural")
          {
              $('#kural').text(messagefromparent.data);
              document.getElementById("kural").value=messagefromparent.data;
              //console.log("Kural :"+messagefromparent.data);
          } else if (messagefromparent.messageType == "Keywords")
          {
           
            document.getElementById("kywrds").value=messagefromparent.data;
            //console.log("Keywords :"+messagefromparent.data);
          } else if (messagefromparent.messageType == "Scope")
          {
           // document.getElementById("scope").value=messagefromparent.data;
            //console.log("scope :"+messagefromparent.data);

          }else if (messagefromparent.messageType == "Class")
          {
           // document.getElementById("class").value=messagefromparent.data;
            //console.log("class :"+messagefromparent.data);
          }else if (messagefromparent.messageType == "Sense")
          {
           // document.getElementById("sense").value=messagefromparent.data;
            //console.log("sense :"+messagefromparent.data);
          }
          else if (messagefromparent.messageType == "ClassScore")
          {
            //document.getElementById("cscore").value=messagefromparent.data;

            //console.log("classScore :"+messagefromparent.data);
            classscore = messagefromparent.data;
          }
          else if (messagefromparent.messageType == "SenseScore")
          {
            //document.getElementById("sscore").value=messagefromparent.data;
            //console.log("senseScore :"+messagefromparent.data);
            sensescore=messagefromparent.data;
          }
          else if (messagefromparent.messageType == "ClassMustbe")
          {
           
            cscoremustbe=messagefromparent.data;
            if(cscoremustbe == null)
             {
               cscoremustbe=1;
             }
            // document.getElementById("cscoremust").value=cscoremustbe;
            //console.log("cscoremust :"+cscoremustbe);
          }
          else if (messagefromparent.messageType == "SenseMustbe")
          {
            sscoremustbe=messagefromparent.data;
            if(sscoremustbe == null)
             {
               sscoremustbe=1;
             }
   
            //document.getElementById("sscoremust").value=sscoremustbe;
            //console.log("sscoremust :"+sscoremustbe);

            populateClasslist(scopeindex,classscore);
          }
          else if (messagefromparent.messageType == "ScopeIndex")
          {
           scopeindex=messagefromparent.data;
            //console.log("scopeindex :"+messagefromparent.data);
          

          }

     }

  function msgtoparent(msgtype,val)
  {
    //console.log("msg to parent");
    var msg = JSON.stringify({messageType:msgtype,data:val});
    Office.context.ui.messageParent(msg);
    //console.log("parent msg yollandı :"+msg);
   
  }

  function populateClasslist(lscopeindex,lclassscore)
     {
       //console.log("class populate started");
      const data = vjson["Scopes"][lscopeindex]["Classifications"];
      //console.log("class length :"+data.length);
      //console.log("classscore: "+lclassscore);
      //console.log("populate classlist started");
       var optClasses=[];
       for(let x=0;x<data.length;x++)
         {
           //console.log("inside loop");
         optClasses[x]="<option disabled data-content="+   "\"" + "<i class='fa fa-square' style='color:"+data[x].Color+"'></i> "+data[x].Name+"\">"+data[x].Id+"</option>";
   
         }
         //console.log("after loop");
         $('#classList').append(optClasses);
         //console.log("after append");
         try{
          $('#classList').selectpicker('refresh');

         } catch (err)
         {
           ////console.log(err);
           msglogger(err);
         }
         //console.log("after refresh class");

   
         $('#classList').selectpicker('val',data[data.length-lclassscore].Id);
 //last one is the default selection
         //console.log(data[data.length-lclassscore].Name);
         enableDropDown(data.length,data);
         PopulateSenseList(sensescore);

     }

 function enableDropDown(datalenght,data)
 {
   var selectmenu;
   selectmenu = document.getElementById("classList").getElementsByTagName("option");
   //console.log("selectmenu lenght: "+selectmenu.length);
   for(let x=datalenght-cscoremustbe;x>-1;x--)
   {
     selectmenu[x].disabled=false;
   }
   //selectmenu[0].disabled=false;
   $('#classList').selectpicker('refresh');
   if(classscore<cscoremustbe)
   {
    $('#classList').selectpicker('val',data[datalenght-cscoremustbe].Id);
    newclassguid = data[datalenght-cscoremustbe].Id;

   } else
   {
    $('#classList').selectpicker('val',data[datalenght-classscore].Id);
    newclassguid = data[datalenght-classscore].Id;

   }

 }
 function PopulateSenseList(lsensescore)
{
  //console.log('Sensitivity List populate started');
  var senseOptions=[];
  const data = vjson["Sensitivities"];
  //console.log("sensitivity json length : " + data.length);
  for(let x=0;x<data.length;x++)
   {
    senseOptions[x]="<option disabled data-content="+   "\"" + "<i class='fa fa-square' style='color:"+data[x].Color+"'></i> "+data[x].Name+"\">"+data[x].Id+"</option>";

   }
$('#sensitivityList').append(senseOptions);
$('#sensitivityList').selectpicker('refresh');
//default selection is 2
//$('#sensitivityList').selectpicker('val',data[lsensescore-1].Id);
  enablesenselist(data.length,data);
  }

  function enablesenselist(datalenght,data)
  {
    var selectmenu;
   selectmenu = document.getElementById("sensitivityList").getElementsByTagName("option");
   for(let x=sscoremustbe-1;x<datalenght;x++)
   {
    selectmenu[x].disabled=false;
   }
   $('#sensitivityList').selectpicker('refresh');
   if(sensescore<sscoremustbe)
   {
    $('#sensitivityList').selectpicker('val',data[sscoremustbe-1].Id);
    newsenseguid=data[sscoremustbe-1].Id;

   } else
   {
   $('#sensitivityList').selectpicker('val',data[sensescore-1].Id);
   newsenseguid=data[sensescore-1].Id;

   }
   //console.log("////Initilization Ended/////") ;
   //console.log("scope index :"+scopeindex);
   //console.log("scope name :"+vjson["Scopes"][scopeindex].Name);
   //console.log("new class guid :"+newclassguid);
   //console.log("new sense guid :"+newsenseguid);

   updateheader();
   /*
   msgtoparent("scope Index",scopeindex);
   msgtoparent("scopename",vjson["Scopes"][scopeindex].Name);
   msgtoparent("classguid",newclassguid);
   msgtoparent("senseguid",newsenseguid);
  */
  }

 function updateheader()
{
  msgtoparent("scope Index",scopeindex);
   msgtoparent("scopename",vjson["Scopes"][scopeindex].Name);
   msgtoparent("classguid",newclassguid);
   msgtoparent("senseguid",newsenseguid);
 
}
















