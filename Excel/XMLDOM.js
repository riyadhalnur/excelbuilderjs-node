'use strict';

var _ = require('underscore');

var XMLDOM = function (ns, rootNodeName) {
  this.documentElement = this.createElement(rootNodeName);
  this.documentElement.setAttribute('xmlns', ns);
};

XMLDOM.prototype.createElement = function (name) {
  return new XMLDOM.XMLNode({
    nodeName: name
  });
};

XMLDOM.prototype.createTextNode = function (text) {
  return new XMLDOM.TextNode(text);
};

XMLDOM.prototype.toString = function () {
  return this.documentElement.toString();
};

XMLDOM.Node = function () {};

XMLDOM.Node.Create = function (config) {
  switch(config.type) {
    case "XML":
      return new XMLDOM.XMLNode(config);
    case "TEXT":
      return new XMLDOM.TextNode(config.nodeValue);
  }
};

XMLDOM.TextNode = function (text) {
  this.nodeValue = text;
};

XMLDOM.TextNode.prototype.toJSON = function () {
  return {
    nodeValue: this.nodeValue,
    type: 'TEXT'
 };
};

XMLDOM.TextNode.prototype.toString = function () {
  return _.escape(this.nodeValue);
};

XMLDOM.XMLNode = function (config) {
  this.nodeName = config.nodeName;
  this.children = [];
  this.nodeValue = config.nodeValue || "";
  this.attributes = {};
  
  if (config.children) {
    for (var i = 0; i < config.children.length; i++) {
      this.appendChild(XMLDOM.Node.Create(config.children[i]));
    }
  }
  
  if (config.attributes) {
    for (var attr in config.attributes) {
      if (config.attributes.hasOwnProperty(attr)) {
        this.setAttribute(attr, config.attributes[attr]);
      }
    }
  }
};

XMLDOM.XMLNode.prototype.toString = function () {
  var string = "<" + this.nodeName;
  var attrs = [];
  for (var attr in this.attributes) {
    if (this.attributes.hasOwnProperty(attr)) {
      attrs.push(attr + "=\""+_.escape(this.attributes[attr])+"\"");
    }
  }
  if (attrs.length > 0){
    string+= " " + attrs.join(" ");
  }

  var childContent = "";
  for (var i = 0, l = this.children.length; i < l; i++) {
    childContent += this.children[i].toString();
  }

  if (childContent){
    string +=  ">" + childContent + "</" + this.nodeName + ">";
  } else {
    string += "/>";
  }
  
  return string;
};

XMLDOM.XMLNode.prototype.toJSON = function () {
  var children = [];
  for (var i = 0, l = this.children.length; i < l; i++) {
    children.push(this.children[i].toJSON());
  }

  return {
    nodeName: this.nodeName,
    children: children,
    nodeValue: this.nodeValue,
    attributes: this.attributes,
    type: "XML"
  };
};

XMLDOM.XMLNode.prototype.setAttribute = function (name, val) {
  if (val === null) {
    delete this.attributes[name];
    delete this[name];
    return;
  }

  this.attributes[name] = val;
  this[name] = val;
};

XMLDOM.XMLNode.prototype.setAttributeNS = function (ns, name, val) {
  this.setAttribute(name, val);
};

XMLDOM.XMLNode.prototype.appendChild = function (child) {
  this.children.push(child);
  this.firstChild = this.children[0];
};

XMLDOM.XMLNode.prototype.cloneNode = function () {
  return new XMLDOM.XMLNode(this.toJSON());
};

module.exports = XMLDOM;
