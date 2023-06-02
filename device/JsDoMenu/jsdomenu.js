/*

  jsDOMenu Version 1.1.3
  Released on 30 December 2003
  jsDOMenu is distributed under the terms of the GNU GPL license
  Refer to readme.txt for more informatiom

*/

// Determine whether the browser is IE5 or IE5.5.
function isIE5() { // Private method
  return (navigator.userAgent.indexOf("MSIE 5") > -1);
}

// Determine whether the browser is IE6.
function isIE6() { // Private method
  return (navigator.userAgent.indexOf("MSIE 6") > -1);
}

// Determine whether the browser is Opera.
function isOpera() { // Private method
  return (navigator.userAgent.indexOf("Opera") > -1);
}

// Determine whether the browser is Safari.
function isSafari() { // Private method
  return ((navigator.userAgent.indexOf("Safari") > -1) && (navigator.userAgent.indexOf("Mac") > -1));
}

/*
Determine the page mode.

0: Quirks mode.
1: Strict mode.
*/
function getPageMode() { // Private method
  if (document.compatMode) {
    switch (document.compatMode) {
      case "BackCompat":
        return 0;
      case "CSS1Compat":
        return 1;
      case "QuirksMode":
        return 0;
    }
  }
  else {
    if (isIE5()) {
      return 0;
    }
    if (isSafari()) {
      return 1;
    }
  }
  return 0;
}

// Alias for document.getElementById().
function getElmId(id) { // Private method
  return document.getElementById(id);
}

// Alias for document.createElement().
function createElm(tagName) { // Private method
  return document.createElement(tagName);
}

// Get the x-coordinate of the mouse position relative to the window.
function getX(e) { // Private method
	if (!e) {
    var e = window.event;
  }
  return e.clientX;
}

// Get the y-coordinate of the mouse position relative to the window.
function getY(e) { // Private method
	if (!e) {
    var e = window.event;
  }
  return e.clientY;
}

// Get the scrollLeft property.
function getScrollLeft() { // Private method
  switch (pageMode) {
    case 0:
      return document.body.scrollLeft;
    case 1:
      if (document.documentElement && (document.documentElement.scrollLeft > 0)) {
        return document.documentElement.scrollLeft;
      }
      else {
        return document.body.scrollLeft;
      }
  }
}

// Get the scrollTop property.
function getScrollTop() { // Private method
  switch (pageMode) {
    case 0:
      return document.body.scrollTop;
    case 1:
      if (document.documentElement && (document.documentElement.scrollTop > 0)) {
        return document.documentElement.scrollTop;
      }
      else {
        return document.body.scrollTop;
      }
  }
}

// Get the clientHeight property.
function getClientHeight() { // Private method
  switch (pageMode) {
    case 0:
      return document.body.clientHeight;
    case 1:
      if ((!isOpera()) && document.documentElement && (document.documentElement.clientHeight > 0)) {
        return document.documentElement.clientHeight;
      }
      else {
        return document.body.clientHeight;
      }
  }
}

// Get the clientWidth property.
function getClientWidth() { // Private method
  switch (pageMode) {
    case 0:
      return document.body.clientWidth;
    case 1:
      if ((!isOpera()) && document.documentElement && (document.documentElement.clientWidth > 0)) {
        return document.documentElement.clientWidth;
      }
      else {
        return document.body.clientWidth;
      }
  }
}

// Adjust pop-up menu position for IE.
function adjustForIE() {
  return ((isIE5() || isIE6()) ? 2 : 0);
}

// Get the left position of the pop-up menu.
function getMainLeftPos(menuObj, x) { // Private method
  if ((x + menuObj.offsetWidth) <= getClientWidth()) {
    return (x - adjustForIE(menuObj));
  }
  else {
    if (x <= getClientWidth()) {
      return (x - menuObj.offsetWidth - adjustForIE(menuObj));
    }
    else {
      return (getClientWidth() - menuObj.offsetWidth);
    }
  }
}

// Get the top position of the pop-up menu.
function getMainTopPos(menuObj, y) { // Private method
  if ((y + menuObj.offsetHeight) <= getClientHeight()) {
    return (y - adjustForIE(menuObj));
  }
  else {
    if (y <= getClientHeight()) {
      return (y - menuObj.offsetHeight - adjustForIE(menuObj));
    }
    else {
      return (getClientHeight() - menuObj.offsetHeight);
    }
  }
}

// Get the left position of the submenu.
function getSubLeftPos(menuObj, x, offset) { // Private method
  if ((x + menuObj.offsetWidth - parseInt(menuObj.style.borderWidth) - 5) <= getClientWidth()) {
    return (x - parseInt(menuObj.style.borderWidth) - 5);
  }
  else {
    if (x <= getClientWidth()) {
      return (x - menuObj.offsetWidth - offset);
    }
    else {
      return (getClientWidth() - menuObj.offsetWidth);
    }
  }
}

// Get the top position of the submenu.
function getSubTopPos(menuObj, y, offset) { // Private method
  if ((y + menuObj.offsetHeight) <= getClientHeight()) {
    return y;
  }
  else {
    if (y <= getClientHeight()) {
      return (y - menuObj.offsetHeight + offset);
    }
    else {
      return (getClientHeight() - menuObj.offsetHeight + offset);
    }
  }
}

// Pop up the main menu.
function popUpMainMenu(e, menuObj) { // Private method
  menuObj.style.left = (getMainLeftPos(menuObj, getX(e)) + getScrollLeft()) + "px";
  menuObj.style.top = (getMainTopPos(menuObj, getY(e)) + getScrollTop()) + "px";
}

// Pop up the submenu.
function popUpSubMenu(parent, menuItem, menuObj) { // Private method
  var x = parent.offsetLeft + parent.offsetWidth - getScrollLeft();
  var y = parent.offsetTop + menuItem.offsetTop - getScrollTop();
  menuObj.style.left = (getSubLeftPos(menuObj, x, menuItem.offsetWidth) + getScrollLeft()) + "px";
  menuObj.style.top = (getSubTopPos(menuObj, y, (menuItem.offsetHeight + (2 * parseInt(menuObj.style.borderWidth)))) + getScrollTop()) + "px";
  if (menuObj.mode == "relative") {
    menuObj.initialLeft = parseInt(menuObj.style.left) - getScrollLeft();
    menuObj.initialTop = parseInt(menuObj.style.top) - getScrollTop();
  }
}

// Show the submenu.
function showSubMenu(menuObj) { // Private method
  popUpSubMenu(menuObj.parent.menuObj, menuObj, menuObj.subMenu.menuObj);
  menuObj.subMenu.menuObj.style.visibility = "visible";
}

// Refresh the menu items.
function refreshMenuItems(menuObj) { // Private method
  for (j = 0; j < menuObj.childNodes.length; j++) {
    if (menuObj.childNodes[j].enabled) {
      menuObj.childNodes[j].className = menuObj.childNodes[j].itemClass;
      if (menuObj.childNodes[j].hasSubMenu) {
        var imgElm = getElmId(menuObj.childNodes[j].id + "Arrow");
        imgElm.src = menuObj.childNodes[j].imgSrc;
      }
    }
  }
}

// Handles mouseover event of the menu item.
function menuItemOver() { // Private method    
  if (this.parent.previousItem) {
    if (this.parent.previousItem.className == this.parent.previousItem.classNameOver) {
        this.parent.previousItem.className = this.parent.previousItem.itemClass;
    }
    if (this.parent.previousItem.hasSubMenu) {
      this.parent.previousItem.className = this.parent.previousItem.itemClass;
      var imgElm = getElmId(this.parent.previousItem.id + "Arrow");
      imgElm.src = this.parent.previousItem.imgSrc;
    }

    for (i = 1; i <= menuCount; i++) {
      var menuObj = getElmId("DOMenu" + i);
      if ((menuObj.level > this.parent.menuObj.level) && (menuObj.mode == this.parent.menuObj.mode)) {
        with (menuObj.style) {
          left = "0px";
          top = "0px";
          visibility = "hidden";
        }
        menuObj.initialLeft = 0;
        menuObj.initialTop = 0;
        //refreshMenuItems(menuObj); //*************************************
      }
    }
  }
  if (this.enabled) {
    this.className = this.classNameOver;
    if (this.hasSubMenu) {
      var imgElm = getElmId(this.id + "Arrow");
      imgElm.src = this.imgSrcOver;
      showSubMenu(this);
    }
  }
  this.parent.previousItem = this;
}

// Handles mouseout event of the menu item.
function menuItemOut() { // Private method
  menuYN="N";
  if (this.enabled) {
    if (!((this.hasSubMenu) && (this.subMenu.menuObj.style.visibility == "visible"))) {
      this.className = this.itemClass;
    }
    if (this.hasSubMenu) {
      var imgElm = getElmId(this.id + "Arrow");
      if (this.subMenu.menuObj.style.visibility == "visible") {
        imgElm.src = this.imgSrcOver;
      }
      else {
        imgElm.src = this.imgSrc;
      }
    }
  }
}

// Handles onclick event of the menu item.
function menuItemClick(e) { // Private method
  if ((this.enabled) && (this.actionOnClick)) {
    var action = this.actionOnClick;
    if (action.indexOf("link:") == 0) {
      location.href = action.substr(5);
    }
    else {
      if (action.indexOf("code:") == 0) {
        eval(action.substr(5));
      }
      else {
        location.href = action;
      }
    }
  }
  if (!e) {
    var e = window.event;
    e.cancelBubble = true;
  }
	if (e.stopPropagation) {
    e.stopPropagation();
  }
  if (this.parent.menuObj.mode == "cursor") {
    hideCursorMenus();
  }
  if ((this.parent.menuObj.mode == "absolute") || (this.parent.menuObj.mode == "relative")) {
    hideVisibleMenus();
  }
}

/*
Set a menu object to be the submenu of another menu object.
Argument:
menuObj        : Required. Menu object that is going to be the submenu.
*/
function setSubMenu(menuObj) { // Public method
  var imgElm = createElm("img");
  imgElm.src = this.imgSrc;
  imgElm.id = this.id + "Arrow";
  with (imgElm.style) {
    position = "absolute";
    right = imgOffset + "px";
    top = Math.floor((this.offsetHeight - imgElm.height) / 2) + "px";
  }
  this.appendChild(imgElm);
  this.hasSubMenu = true;
  this.subMenu = menuObj;
  this.imgElm = imgElm
  this.setImgSrc = function(imgSrc) { // Public method
    if (this.imgElm) {
      this.imgSrc = imgSrc;
      this.imgElm.src = this.imgSrc
      preloadImg(imgSrc);
    }
  }
  this.setImgSrcOver = function(imgSrcOver) { // Public method
    if (this.imgElm) {
      this.imgSrcOver = imgSrcOver;
      preloadImg(imgSrcOver);
    }
  }
  this.setImgOffset = function(offset) { // Public method
    if (this.imgElm) {
      this.imgElm.style.right = offset + "px";
    }
  }
  menuObj.menuObj.style.zIndex = this.parent.menuObj.level + 1;
  menuObj.menuObj.level = this.parent.menuObj.level + 1;
}

/*
Adds a new menu item to the menu.
Argument:
menuItemObj        : Required. Menu item object that is going to be added to the menu object.
*/
function addMenuItem(menuItemObj) { // Public method
  if (menuItemObj.displayText == "-") {
    var hrElm = createElm("hr");
    var itemElm = createElm("div");
    itemElm.appendChild(hrElm);
    itemElm.id = menuItemObj.id;
    if (menuItemObj.className.length > 0) {
      itemElm.sepClass = menuItemObj.className;
    }
    else {
      itemElm.sepClass = menuItemObj.sepClass;
    }
    itemElm.className = itemElm.sepClass;
    this.menuObj.appendChild(itemElm);
    itemElm.parent = this;
    itemElm.setClassName = function(className) { // Public method
      this.sepClass = className;
      this.className = this.sepClass;
    }
    itemElm.onclick = function(e) { // Private method
      if (!e) {
        var e = window.event;
        e.cancelBubble = true;
      }
      if (e.stopPropagation) {
        e.stopPropagation();
      }
    }
    itemElm.onmouseover = menuItemOver;
    if (menuItemObj.itemName.length > 0) {
      this.items[menuItemObj.itemName] = itemElm;
    }
    else {
      this.items[this.items.length] = itemElm;
    }
  }
  else {
    var itemElm = createElm("div");
    itemElm.id = menuItemObj.id;
    itemElm.actionOnClick = menuItemObj.actionOnClick;
    itemElm.enabled = menuItemObj.enabled;
    itemElm.itemClass = menuItemObj.className;
    itemElm.classNameOver = menuItemObj.classNameOver;
    itemElm.className = itemElm.itemClass;
    itemElm.hasSubMenu = false;
    itemElm.subMenu = null;
    itemElm.imgSrc = imgSrc;
    itemElm.imgSrcOver = imgSrcOver;
    this.menuObj.appendChild(itemElm);
    var textNode = document.createTextNode(menuItemObj.displayText);
    itemElm.appendChild(textNode);
    itemElm.parent = this;
    itemElm.setClassName = function(className) { // Public method
      this.itemClass = className;
      this.className = this.itemClass;
    }
    itemElm.setClassNameOver = function(classNameOver) { // Public method
      this.classNameOver = classNameOver;
    }
    itemElm.setDisplayText = function(text) { // Public method
      this.firstChild.nodeValue = text;
    }
    itemElm.setSubMenu = setSubMenu;
    itemElm.onclick = menuItemClick;
    itemElm.onmouseover = menuItemOver;
    itemElm.onmouseout = menuItemOut;
    if (menuItemObj.itemName.length > 0) {
      this.items[menuItemObj.itemName] = itemElm;
    }
    else {
      this.items[this.items.length] = itemElm;
    }
  }
}

/*
Creates a new menu item object.
Arguments:
displayText        : Required. String that specifies the text to be displayed on the menu item. If 
                     displayText is "-", a menu separator will be created instead.
itemName           : Optional. String that specifies the name of the menu item. Defaults to "".
actionOnClick      : Optional. String that specifies the action to be done when the menu item is 
                     being clicked. Defaults to "" (no actions).
enabled            : Optional. Boolean that specifies whether the menu item is enabled/disabled. 
                     Defaults to true.
className          : Optional. String that specifies the CSS style selector for the menu item. 
                     Defaults to "jsdomenuitem".
classNameOver      : Optional. String that specifies the CSS style selector for the menu item when the 
                     mouse is over it. Defaults to "jsdomenuitemover".
*/
function menuItem() { // Public method
 this.displayText = arguments[0];
 if (this.displayText == "-") {
    this.id = "Sep" + (++ sepCount);
    this.className = sepClass;
  }
  else {
    this.id = "Item" + (++ itemCount);
    this.className = itemClass;
  }
  this.sepClass = sepClass;
  this.itemName = "";
  this.actionOnClick = "";
  this.enabled = true;
  this.classNameOver = itemClassOver;
  var length = arguments.length;
  if ((length > 1) && (arguments[1].length > 0)) {
    this.itemName = arguments[1];
  }
  if ((length > 2) && (arguments[2].length > 0)) {
    this.actionOnClick = arguments[2];
  }
  if ((length > 3) && (typeof(arguments[3]) == "boolean")) {
    this.enabled = arguments[3];
  }
  if ((length > 4) && (arguments[4].length > 0)) {
    if (arguments[4] == "-") {
      this.className = arguments[4];
      this.sepClass = arguments[4];
    }
    else {
      this.className = arguments[4];
    }
  }
  if ((length > 5) && (arguments[5].length > 0)) {
    this.classNameOver = arguments[5];
  }
}

/*
Creates a new menu object.
Arguments:
width              : Required. Integer that specifies the width of the menu.
className          : Optional. String that specifies the CSS style selector for the menu. Defaults 
                     to "jsdomenudiv".
mode               : Optional. String that specifies the mode of the menu. Defaults to "cursor".
alwaysVisible      : Optional. Boolean that specifies whether the menu is always visible. Defaults to 
                     false.
*/
function jsDOMenu() { // Public method
  this.items = new Array();
  var menuObj = createElm("div");
  menuObj.id = "DOMenu" + (++menuCount);
  menuObj.level = 10;
  menuObj.previousItem = null;
  menuObj.allExceptFilter = allExceptFilter;
  menuObj.noneExceptFilter = noneExceptFilter;
  menuObj.className = menuClass
  menuObj.mode = "cursor";
  menuObj.alwaysVisible = false;
  menuObj.initialLeft = 0;
  menuObj.initialTop = 0;
  var length = arguments.length;
  if ((length > 1) && (arguments[1].length > 0)) {
      menuObj.className = arguments[1];
    }
  if ((length > 2) && (arguments[2].length > 0)) {
    switch (arguments[2]) {
      case "absolute":
        menuObj.mode = "absolute";
        break;
      case "relative":
        menuObj.mode = "relative";
        break;
    }
  }  
  if ((length > 3) && (typeof(arguments[3]) == "boolean")) {
    menuObj.alwaysVisible = arguments[3];
  }
  with (menuObj.style) {
    width = arguments[0] + "px";
    left = "0px";
    top = "0px";
    borderWidth = menuBorderWidth + "px";
  }
  document.body.appendChild(menuObj);
  this.menuObj = menuObj;
  this.addMenuItem = addMenuItem;
  this.setClassName = function(className) { // Public method
    this.menuObj.className = className;
  }
  this.setMode = function(mode) { // Public method
    this.menuObj.mode = mode;
  }
  this.setAlwaysVisible = function(alwaysVisible) { // Public method
    this.menuObj.alwaysVisible = alwaysVisible;
  }

  this.show = function() { // Public method
    this.menuObj.style.visibility = "visible";
  }
  this.hide = function() { // Public method
    this.menuObj.style.visibility = "hidden";
    if (this.menuObj.mode == "cursor") {
      with (this.menuObj.style) {
        left = "0px";
        top = "0px";
      }
      menuObj.initialLeft = 0;
      menuObj.initialTop = 0;
    }
  }
  this.setX = function(x) { // Public method
    this.menuObj.initialLeft = x;
    this.menuObj.style.left = x + "px";
  }
  this.setY = function(y) { // Public method
    this.menuObj.initialTop = y;
    this.menuObj.style.top = y + "px";
  }
  this.moveTo = function(x, y) { // Public method
    this.menuObj.initialLeft = x;
    this.menuObj.initialTop = y;
    this.menuObj.style.left = x + "px";
    this.menuObj.style.top = y + "px";
  }
  this.moveBy = function(x, y) { // Public method
    var left = parseInt(this.menuObj.style.left);
    var top = parseInt(this.menuObj.style.top);
    this.menuObj.initialLeft = left + x;
    this.menuObj.initialTop = top + y;
    this.menuObj.style.left = (left + x) + "px";
    this.menuObj.style.top = (top + y) + "px";
  }
  this.setAllExceptFilter = function(filter) { // Public method
    this.menuObj.allExceptFilter = filter;
    this.menuObj.noneExceptFilter = new Array();
  }
  this.setNoneExceptFilter = function(filter) { // Public method
    this.menuObj.noneExceptFilter = filter;
    this.menuObj.allExceptFilter = new Array();
  }
  this.setBorderWidth = function(width) { // Public method
    this.menuObj.style.borderWidth = width + "px";
  }
}


// Find out whether any of the tag name/tag id pair in the filter matches the tagName/tagId pair.
function findMatch(tagName, tagId, filter) { // Private method
  for (i = 0; i < filter.length; i++) {
    var temp = filter[i].toLowerCase().split(".");
    if (((temp[0] == "*") && (temp[1] == "*")) || 
        ((temp[0] == "*") && (temp[1] == tagId)) ||
        ((temp[0] == tagName) && (temp[1] == "*")) ||
        ((temp[0] == tagName) && (temp[1] == tagId))) {
      return true;
    }
  }
  return false;
}

//Determine whether to show or hide the menu.
function canShowMenu(tagName, tagId, allExcept, noneExcept) { // Private method
  if (allExcept.length > 0) {
    return (!findMatch(tagName.toLowerCase(), tagId.toLowerCase(), allExcept));
  }
  else {
    if (noneExcept.length > 0) {
      return findMatch(tagName.toLowerCase(), tagId.toLowerCase(), noneExcept);
    }
    else {
      return true;
    }
  }
}

// Shows/Hides the pop-up menu.
function activatePopUpMenu(e) { // Private method
  if (!popUpMenu) {
    return;
  }
  var state = popUpMenu.menuObj.style.visibility;
  if (state == "visible") {
    for (i = 1; i <= menuCount; i++) {
      var menuObj = getElmId("DOMenu" + i);
      if (menuObj.mode == "cursor") {
        with (menuObj.style) {
          left = "0px";
          top = "0px";
          visibility = "hidden";
        }
        menuObj.initialLeft = 0;
        menuObj.initialTop = 0;
        refreshMenuItems(menuObj);
      }
    }
  }
  else {
    if (!e) {
      var e = window.event;
    }
    var tg = (e.target) ? e.target : e.srcElement;
    var popUpMenuObj = popUpMenu.menuObj;
    if (canShowMenu(tg.tagName, tg.id, popUpMenuObj.allExceptFilter, popUpMenuObj.noneExceptFilter)) {
      popUpMainMenu(e, popUpMenuObj);
      setTimeout("popUpMenu.menuObj.style.visibility = 'visible'", menuDelay);
    }
  }
}

/*
Specifies how the pop-up menu shows/hide.
Arguments:
showValue          : Required. Integer that specifies how the menu shows.
hideValue          : Optional. Integer that specifies how the menu hides. If not specified, the 
                     menu shows/hides in the same manner.

0: Shows/Hides the menu by left click only.
1: Shows/Hides the menu by right click only.
2: Shows/Hides the menu by left or right click.
*/
function activatePopUpMenuBy() { // Public method
  showValue = ((typeof(arguments[0]) == "number") && (arguments[0] > -1)) ? arguments[0] : 0;
  if (arguments.length > 1) {
    hideValue = ((typeof(arguments[1]) == "number") && (arguments[1] > -1)) ? arguments[1] : 0;
  }
  else {
    hideValue = showValue;
  }
  if ((showValue == 1) || (showValue == 2) || (hideValue == 1) || (hideValue == 2)) {
    document.oncontextmenu = rightClickHandler;
  }
}

// Hiding all menus except those with alwaysVisible = true.
function hideAllMenus() { // Public method
  for (i = 1; i <= menuCount; i++) {
    var menuObj = getElmId("DOMenu" + i);
    if (!menuObj.alwaysVisible) {
      menuObj.style.visibility = "hidden";
      if (menuObj.mode == "cursor") {
        with (menuObj.style) {
          left = "0px";
          top = "0px";
        }
        menuObj.initialLeft = 0;
        menuObj.initialTop = 0;
      }
    }
    refreshMenuItems(menuObj);
  }
}

// Hiding all menus with mode = "cursor", except those with alwaysVisible = true.
function hideCursorMenus() { // Public method
  for (i = 1; i <= menuCount; i++) {
    var menuObj = getElmId("DOMenu" + i);
    if ((menuObj.mode == "cursor") && (!menuObj.alwaysVisible)) {
      with (menuObj.style) {
        left = "0px";
        top = "0px";
        visibility = "hidden";
      }
      menuObj.initialLeft = 0;
      menuObj.initialTop = 0;
    }
    if (menuObj.mode == "cursor") {
      refreshMenuItems(menuObj);
    }
  }
}

// Hiding all menus with mode = "absolute" or mode = "relative", except those with alwaysVisible = true.
function hideVisibleMenus() { // Public method
  for (i = 1; i <= menuCount; i++) {
    var menuObj = getElmId("DOMenu" + i);
    if (((menuObj.mode == "absolute") || (menuObj.mode == "relative")) && (!menuObj.alwaysVisible)) {
      menuObj.style.visibility = "hidden";
    }
    if ((menuObj.mode == "absolute") || (menuObj.mode == "relative")) {
      refreshMenuItems(menuObj);
    }
  }
}

// Handles left click event.
function leftClickHandler(e) { // Private method
  if ((getX(e) > getClientWidth()) || (getY(e) > getClientHeight())) {
    return;
  }
  if (!e) {
    var e = window.event;
  }
	if ((e.button) && (e.button == 2)) {
    return;	
  }
  hideVisibleMenus();
  if (popUpMenu) {
    var state = popUpMenu.menuObj.style.visibility;
    if ((state == "visible") && ((hideValue == 0) || (hideValue == 2))) {
      activatePopUpMenu(e);
    }
    if (((state == "hidden") || (state == "")) && ((showValue == 0) || (showValue == 2))) {
      activatePopUpMenu(e);
    }
  }
}

// Handles right click event.
function rightClickHandler(e) { // Private method
  if ((getX(e) > getClientWidth()) || (getY(e) > getClientHeight())) {
    return;
  }
  if (popUpMenu) {
    var state = popUpMenu.menuObj.style.visibility;
    if ((state == "visible") && ((hideValue == 1) || (hideValue == 2))) {
      activatePopUpMenu(e);
      return false;
    }
    if (((state == "hidden") || (state == "")) && ((showValue == 1) || (showValue == 2))) {
      activatePopUpMenu(e);
      return false;
    }
  }
}

// Handle scroll event.
function scrollHandler() { // Private method
  for (i = 1; i <= menuCount; i++) {
      var menuObj = getElmId("DOMenu" + i);
      if (menuObj.mode == "relative") {
        with (menuObj.style) {
          left = (menuObj.initialLeft + getScrollLeft()) + "px";
          top = (menuObj.initialTop + getScrollTop()) + "px";
        }
      }
  }
}

/*
Set which menu object is the pop-up menu
Argument:
menuObj            : Required. Menu object that specifies the pop-up menu.
*/
function setPopUpMenu(menuObj) { // Public method
  popUpMenu = menuObj;
}

/* Preload the image.
Argument:
imgSrc             : Required. String that specifies the image source to preload.
*/
function preloadImg(imgSrc) { // Public method
  var img = new Image();
  img.src = imgSrc;
}

// Checks browser compatibility and creates the menus.
function initjsDOMenu() { // Public method
  if (document.createElement && document.getElementById) {
    createjsDOMenu();
  }
}

if (typeof(allExceptFilter) == "undefined") {
  var allExceptFilter = new Array("A.*", 
                                  "BUTTON.*", 
                                  "IMG.*", 
                                  "INPUT.*", 
                                  "OBJECT.*", 
                                  "OPTION.*", 
                                  "SELECT.*", 
                                  "TEXTAREA.*"); // Public property
}

if (typeof(noneExceptFilter) == "undefined") {
  var noneExceptFilter = new Array(); // Public property
}

if (typeof(menuClass) == "undefined") {
  var menuClass = "jsdomenudiv"; // Public property
}

if (typeof(itemClass) == "undefined") {
  var itemClass = "jsdomenuitem"; // Public property
}

if (typeof(itemClassOver) == "undefined") {
  var itemClassOver = "jsdomenuitemover"; // Public property
}

if (typeof(sepClass) == "undefined") {
  var sepClass = "jsdomenusep"; // Public property
}

if (typeof(imgSrc) == "undefined") {
  var imgSrc = "JsDoMenu/arrow.png"; // Public property
}

if (typeof(imgSrcOver) == "undefined") {
  var imgSrcOver = "JsDoMenu/arrow_o.png"; // Public property
}

if (typeof(imgOffset) == "undefined") {
  var imgOffset = 8; // Public property
}

if (typeof(menuBorderWidth) == "undefined") {
  var menuBorderWidth = 2; // Public property
}

if (typeof(menuDelay) == "undefined") {
  var menuDelay = 0; // Public property
}

var menuYN = "N";
var menuCount = 0; // Private property
var itemCount = 0; // Private property
var sepCount = 0; // Private property
var popUpMenu = null; // Private property
var showValue = 0; // Private property
var hideValue = 0; // Private property
var pageMode = getPageMode(); // Private property
preloadImg(imgSrc);
preloadImg(imgSrcOver);
document.onclick = leftClickHandler;
window.onscroll = scrollHandler;

var menutimeoutID = 0;
function HideTimeOutMenus() { 
  if(menutimeoutID != 0) {window.clearTimeout(menutimeoutID); menutimeoutID=0;}
  menutimeoutID=window.setTimeout("HidingAllMenus()",3500);
}

function MenusOut() { 
  menuYN="N";
  if(menutimeoutID != 0) {window.clearTimeout(menutimeoutID); menutimeoutID=0;}
  menutimeoutID=window.setTimeout("HidingAllMenus()",3000);
}

// Hiding all menus
function HidingAllMenus() { // Public method
  if ( menuYN=="N" )
  { for (i = 1; i <= menuCount; i++)
    { var menuObj = getElmId("DOMenu" + i);
      menuObj.style.visibility = "hidden";
    }    
  }
}
