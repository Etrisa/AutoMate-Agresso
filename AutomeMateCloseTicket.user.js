// ==UserScript==
// @name         AutomeMateCloseTicket
// @namespace    Cube
// @version      1.0
// @description  @AMFA# AutoMateFuckAgresso, part of an automatic system to not do agresso manually.
// @author       Johan Samuelsson
// @match        https://cube.advania.se/*
// @grant        none
// @run-at       document-end
// ==/UserScript==

function waitFor(waitFor) {
    return new Promise((resolve, reject) => {
       function perform() {
           if (waitFor()) {
               resolve();
           } else {
               setTimeout(perform, 500);
           }
       }
       setTimeout(perform, 500);
    });
}

var labelsDescription = " \
    <table> \
    <tbody>  \
    <tr>  \
        <td>Inom Avtal</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">Ja</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">Nej</td>  \
    </tr>  \
    <tr>  \
        <td>Worksite inom avtal</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WsIn</td>  \
    </tr>  \
    <tr>  \
        <td>Worksite utanför avtal</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WsOut</td>  \
    </tr>  \
    <tr>  \
        <td>Driftavtal</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">Drift</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal WSA Stockholm</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WSASth</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal WSA Skåne</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WSASkn</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal WSA Göteborg/Borås</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WSAGbg</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal WSA Jönköping</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WSAJkp</td>  \
    </tr>  \
    <tr>  \
        <td>Utanför avtal WSA Storbyrå/Hela</td>  \
        <td>=</td>  \
        <td style=\"text-align:left\">WSAGlobal</td>  \
    </tr>  \
    </tbody>  \
    </table>"

async function fix() {
  await waitFor(() => document.querySelector("#multiple-input-journal-entry > div:nth-child(1) > label > span.label-text.ng-binding"));
    var linebreak = document.createElement("br");

    var btn = document.createElement("BUTTON");
    btn.innerHTML = "AutoMate Close";
    btn.onclick = WritePrivateNote;

    var text = document.createElement("div");
    text.innerHTML = labelsDescription;

    var element = document.querySelector("#multiple-input-journal-entry > div:nth-child(1) > label > span.label-text.ng-binding");
    element.appendChild(linebreak);
    element.appendChild(btn);
    element.appendChild(text);
}
fix();

function WritePrivateNote(){
    //Get company name, this will only be used for WSA.
    /*var Company = null
    var INCCompany = document.querySelector("#sys_display\\.incident\\.company");
    var REQCompany = document.querySelector("#sys_display\\.sc_req_item\\.company");

    if(INCCompany){
        Company = INCCompany.value;
       }else if(REQCompany){
       Company = REQCompany.value;
       }else{
       alert("Shit's fucked, be johan fixa skiten. (No Company found.)")
       }*/// Nevermind this will not be used at all

    //Get the name of the person who sent in the ticket.
    var Name = null;
    var CallerId = document.querySelector("#sys_display\\.incident\\.caller_id");
    var RequestedFor = document.querySelector("#sys_display\\.sc_req_item\\.request\\.requested_for");

    if(CallerId){
        Name = CallerId.value;
    }else if(RequestedFor){
        Name = RequestedFor.value;
    }else{
        alert("Shit's fucked, be johan fixa skiten. (No caller ID or Requested For is found)");
        return;
    }

    //Get the short description of the ticket
    var ShortDescription = null;
    var INCShortDescription = document.querySelector("#incident\\.short_description");
    var REQShortDescription = document.querySelector("#sc_req_item\\.short_description");

    if(INCShortDescription){
        ShortDescription = INCShortDescription.value;
    }else if(REQShortDescription){
        ShortDescription = REQShortDescription.value;
    }else{
        alert("Shit's fucked, be johan fixa skiten. (short description not found)");
        return;
    }

    //Write private note
    document.querySelector("#activity-stream-u_private_notes-textarea").value = "Name: " + Name +
        "\nShort description: " + ShortDescription +
        "\nClose notes: " +
        "\nAvtal: Ja" +
        "\nTime (h.m): 1.0";
    return false;
}

//Ticketnumber: INC123456 (Gets this from export)
//Name: Anna Andersson
//Short description: Problem med skrivare
//Close notes: Ändrar standardskrivare
//Inom Avtal (Ja/Nej): Ja
//Time (format h.m): 1.5