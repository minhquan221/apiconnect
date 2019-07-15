var express = require('express');
//var bodyParser = require('body-parser');
var app = express();
var router = require("./ServiceApp");
//var SyncData = require("./syncdata").SyncData;
var port = 1337;
var http = require('http').createServer();
var io = require('socket.io').listen(http);
var mongoData = require("./mongosocial").mongoData;
var history = [];
var Guests = [];
var Users = [];
app.use("/", router);


app.listen(4830, function() {
    console.log("Server NodeJS Listening on port 4830");
    //SyncData.Synchornize();
});

http.listen(port, function() {
    console.log('Socket IO Server listening on port ' + port);
});

io.sockets.on('connection', function(socket) {
    //socket.join('Guest');

    var socketcurrentid = '';
    //clients[socket.id] = socket;
    socketcurrentid = socket.id;
    var UserNameClient = '';
    //clients.push({ username: UserNameClient, socketId: socketcurrentid });
    //io.to(socketcurrentid).emit('loadhistory', history);
    //io.to(socketcurrentid).emit('loadpost');
    //console.log(clients);
    console.log('A new Client has connected: ' + socket.id);
    // var countGuest = 0;
    // var countUser = 0;
    // clients.forEach(function(item, index, object) {
    //     if (item.username == '') {
    //         countGuest++;
    //     }
    //     if (item.username != '') {
    //         countUser++;
    //     }
    // });
    // io.emit('notifyNumberOfUser', countUser);
    // io.emit('notifyNumberOfGuest', countGuest);
    io.to(socketcurrentid).emit('identify');

    socket.on('firstconnect', function() {
        io.to(socketcurrentid).emit('confirmconnect', socketcurrentid);
        if (UserNameClient == '') {
            Guests.push({ username: UserNameClient, socketId: socketcurrentid });
        } else {
            Users.push({ username: UserNameClient, socketId: socketcurrentid });
        }
        //clients.push({ username: UserNameClient, socketId: socketcurrentid });
    });
    socket.on('userReload', function(data) {
        console.log('Reload User ' + data);
        var countIn = 0;
        clients.forEach(function(item, index, object) {
            if (item.socketId == data) {
                countIn++;
            }
        });
        console.log('Count ' + countIn);
        if (countIn > 0) {
            io.to(socketcurrentid).emit('reloadconfirm', data);
            socketcurrentid = data;
            //socket.id = socketcurrentid;
            socket.join(socketcurrentid);
        } else {
            io.to(socketcurrentid).emit('renewrequired', socketcurrentid);
            clients.push({ username: UserNameClient, socketId: socketcurrentid });
        }
    });

    socket.on('error', function(error) {
        // send a private message to the socket with the given id
        console.log('Error at ' + error);
    });
    socket.on('loadpost', async function() {
        var query = {};
        var postlist = await mongoData.GetPostList(query);
        io.to(socketcurrentid).emit('sendpost', postlist);
    })
    socket.on('received', function(msg) {
        console.log("Received message " + msg);
        var json = JSON.parse(msg);
        json['time'] = new Date().getHours().toString() + ":" + new Date().getMinutes().toString() + ":" + new Date().getSeconds().toString() + " " + new Date().getDay().toString() + "/" + new Date().getMonth().toString() + "/" + new Date().getFullYear().toString();
        io.emit('message', JSON.stringify(json));
        history.push(json);
    })
    socket.on('broadcastcenter', function() {
        //do something
    });
    socket.on('message', function(msg) {
        console.log("Received message " + msg);
        var json = JSON.parse(msg);
        json['time'] = new Date().getHours().toString() + ":" + new Date().getMinutes().toString() + ":" + new Date().getSeconds().toString() + " " + new Date().getDay().toString() + "/" + new Date().getMonth().toString() + "/" + new Date().getFullYear().toString();
        io.emit('message', JSON.stringify(json));
        history.push(json);
    });
    socket.on('logged', function(data) {
        UserNameClient = data.UserName;
        clients.forEach(function(item, index, object) {
            if (item.socketId == socketcurrentid) {
                object.splice(index, 1);
            }
        });
        clients.push({ username: UserNameClient, socketId: socketcurrentid });
        var countGuest = 0;
        var countUser = 0;
        clients.forEach(function(item, index, object) {
            if (item.username == '') {
                countGuest++;
            }
            if (item.username != '') {
                countUser++;
            }
        });
        io.emit('notifyNumberOfUser', countUser);
        io.emit('notifyNumberOfGuest', countGuest);
    });
    socket.on('disconnect', function() {
        //setTimeout(function() {
        // clients.forEach(function(item, index, object) {
        //     if (item.socketId == socketcurrentid) {
        //         countIn++;
        //         object.splice(index, 1);
        //     }
        // });
        // console.log('User from socket ' + socketcurrentid + ' disconnected');
        // var countGuest = 0;
        // var countUser = 0;
        // clients.forEach(function(item, index, object) {
        //     if (item.username == '') {
        //         countGuest++;
        //     }
        //     if (item.username != '') {
        //         countUser++;
        //     }
        // });
        // io.emit('notifyNumberOfUser', countUser);
        // io.emit('notifyNumberOfGuest', countGuest);
        //}, 10000);
    });
});