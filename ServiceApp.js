var express = require('express');
var bodyParser = require('body-parser');
var app = express();
var bodyParser = require('body-parser');
var express = require("express");
var router = new express.Router();
//var logger = require("./logger").Logger;
var mongoData = require("./mongosocial").mongoData;

// router.use(function timeLog(req, res, next) {
// // this is an example of how you would call our new logging system to log an info message
// logger.info("Test Message");
// next();
// });
router.use(bodyParser.json());
router.post("/register", async function(request, res) {
    var data = request.body;
    var resultdata = await mongoData.UserRegister(data);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})
router.post("/login", async function(request, res) {
    var data = request.body;
    var resultdata = await mongoData.UserLogin(data);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})
router.post("/post", async function(request, res) {
    var data = request.body;
    data.CreatedDate = new Date().valueOf();
    var resultdata = await mongoData.PostContent(data);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})

router.post("/comment", async function(request, res) {
    var data = request.body;
    data.CreatedDate = new Date().valueOf();
    var resultdata = await mongoData.CommentPost(data);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})

router.post("/getcomment", async function(request, res) {
    var PageIndex = request.body.PageIndex;
    var PageSize = request.body.PageSize;
    var IdPost = request.body.IdPost;
    var resultdata = await mongoData.GetCommentList(IdPost, PageIndex, PageSize);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})

router.post("/getlistpost", async function(request, res) {
    var PageIndex = request.body.PageIndex;
    var PageSize = request.body.PageSize;
    var query = request.body.query;
    var resultdata = await mongoData.GetPostList(query, PageIndex, PageSize);
    if (resultdata.error == 0) {
        res.send(resultdata.data);
    } else {
        res.send(resultdata.msg);
    }
})

module.exports = router;