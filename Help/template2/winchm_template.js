function d_onresize(){
if (window.navigator.userAgent.indexOf("MSIE")>=1){
document.getElementById('winchm_template_content').style.pixelWidth = document.body.offsetWidth - 3;
document.getElementById('winchm_template_content').style.pixelHeight = document.body.offsetHeight - document.getElementById('winchm_template_top').offsetHeight - 4;
}
document.getElementById('winchm_template_content').style.top = document.getElementById('winchm_template_top').offsetHeight + 'px';
}

function d_onbeforeprint(){
document.getElementById('winchm_template_content').style.width = 'auto';
document.getElementById('winchm_template_content').style.height = 'auto';
}

function d_onafterprint(){
d_onresize();
}

window.onload = d_onresize;
window.onresize = d_onresize;
window.onbeforeprint = d_onbeforeprint;
window.onafterprint = d_onafterprint;