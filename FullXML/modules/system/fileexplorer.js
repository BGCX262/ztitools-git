// object for the preview picture. used to get the size and resize to <=150px
var _img;


// show options in the information box (top left of file list)
function showOptions(typ, nm) { 
  PreviewImage(nm);
 	if (window.opener)
	{
		fileinfos.innerHTML += '<a href="javascript:SelectThisImage(\''+nm+'\');">Select this picture.</a>';
	}
}

// user has clicked on the "select this picture" link
function SelectThisImage(p_sPath)
{
	window.opener.setImage( p_sPath ) ;
	window.close() ;
}

function PreviewImage(p_sPath)
{
	//hide old picture
	preview.style.display = "none";
	
	//load new picture
	_img = new Image();
	_img.src = p_sPath;	
	_img.onload = imgPreview_Load;	
}

function imgPreview_Load()
{
	// resize
	if (_img.width>150)
		_img.width = 150;
	
	preview.innerHTML = '<img src="' + _img.src + '" width=' + _img.width + '>';
	//preview.innerHTML = '<img id=_img width=' + _img.width + '>';
	_img = null; 
	
	preview.style.display = "block";
}