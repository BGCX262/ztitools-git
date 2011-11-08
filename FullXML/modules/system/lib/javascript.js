/*

Vous ne devez pas suprimer cet en-tête
fichier javascript pour créer un menu déroulant automatiquement

nombre de niveaux non limité

Auteur : bieler batiste
Blog : www.magnin-sante.ch/journal

send me a mail : http://www.magnin-sante.ch/journal/?p=mailto.php&amp;m=gbefoAbmufso/psh

*/
       
function initmenu()
{

var menu = document.getElementById('hmenu');
var lis = menu.getElementsByTagName('li');
var uls = menu.getElementsByTagName('ul');

    if ( !browser.isOpera || browser.versionMajor>=7 )
    {
        menu.className="hmenu"; /* attach the style to the menu */
        
        for ( var i=0; i<lis.length; i++ )
        {
            var ul = lis.item(i).getElementsByTagName('ul');
            
            if ( ul.item(0) )
            {
                if ( browser.isIE ) /* for Internet Explorer */
                {
                    lis.item(i).onmouseover = visible;
                    lis.item(i).onmouseout = hidden;
                    lis.item(i).onkeyup = visible;
                }
                else /* for Browser */
                {
                    lis.item(i).addEventListener("mouseover",visible,true);
                    lis.item(i).addEventListener("mouseout",hidden,true);
                    lis.item(i).addEventListener("blur",hidden,true);
                    lis.item(i).addEventListener("focus",visible,true);
                }
            }
        }
    }
}
    
function hiddenUl( ul )
{
    var uls = ul.getElementsByTagName('ul');
    for ( var i=0; i<uls.length; i++ )
    {
        uls.item(i).style.visibility = "hidden";
    }
    ul.style.visibility = "hidden";
} 
    
function hidden(){
    var ul = this.getElementsByTagName('ul');
    ul.item(0).style.visibility = "hidden";
    }
    
function visible(){
    var ul = this.getElementsByTagName('ul');
    ul.item(0).style.visibility = "visible";
    }
    
    