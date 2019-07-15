var mongoClient = require('mongodb').MongoClient;
var url = "mongodb://127.0.0.1:27017/";
var dbName = "SocialDB";
var options = {
    useNewUrlParser: true
};
var mongoData = (exports.mongoData = {});
//async function(query, collection, pageindex, pagesize)
mongoData.UserRegister = async function(objectdata) {
    var returnObject = mongoData.Insert(objectdata, 'User');
    return returnObject;
}
mongoData.UserLogin = async function(UserLogin) {
    var returnObject = await mongoData.GetOne(UserLogin, "User");
    return returnObject;
};

mongoData.PostContent = async function(data) {
    var returnObject = await mongoData.Insert(data, "Post");
    return returnObject;
};

mongoData.CommentPost = async function(data) {
    var returnObject = await mongoData.Insert(data, "Comment");
    return returnObject;
};

mongoData.GetCommentList = async function(IdPost, pageindex, pagesize) {
    var query = {};
    query["IdPost"] = IdPost;
    var returnObject = await mongoData.QueryData(query, "Comment", pageindex, pagesize);
    return returnObject;
}

mongoData.GetPostList = async function(query, pageindex = 1, pagesize = 20) {
    var query = {};
    var returnObject = await mongoData.QueryData(query, "Post", pageindex, pagesize);
    return returnObject;
}

mongoData.GetOne = async function(data, tableName) {
    var returnObject = { error: 0, msg: '', TotalRows: 0, data: null };
    var query = {};
    for (var i = 0; i < Object.keys(data).length; i++) {
        query[Object.keys(data)[i]] = data[Object.keys(data)[i]];
    }
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata = await dbo.collection(tableName).findOne(query);
        if (resultdata != undefined) {
            returnObject.TotalRows = 1;
            returnObject.data = resultdata;
        }
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.WriteCache = async function(objectMongo, tableName) {
    var returnObject = { error: 0, msg: '', data: null };
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        await dbo.collection(tableName).deleteOne({
            "_id": objectMongo._id
        });
        await dbo.collection(tableName).insertOne(objectMongo);
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.ClearCache = async function() {
    var returnObject = { error: 0, msg: '', data: null };
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata = await dbo.collection(CacheTableName).deleteMany({ "_id": { $ne: null } });
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.QueryData = async function(query, collection, pageindex = 1, pagesize = null) {
    var returnObject = { error: 0, msg: '', TotalRows: 0, data: null };
    var skipdata = 0;
    if (pagesize != null)
        skipdata = (parseInt(pageindex) - 1) * parseInt(pagesize);
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata;
        if (pagesize != null) {
            resultdata = await dbo.collection(collection).find(query).skip(skipdata).limit(parseInt(pagesize)).toArray();
        } else {
            resultdata = await dbo.collection(collection).find(query).toArray();
        }
        if (resultdata != undefined && resultdata.length > 0) {
            var totalRow = await dbo.collection(collection).countDocuments(query);
            returnObject.TotalRows = totalRow;
            returnObject.data = resultdata;
        }
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.GetAll = async function(data, tableName) {
    var returnObject = { error: 0, msg: '', TotalRows: 0, data: null };
    var query = {};
    for (var i = 0; i < Object.keys(data).length; i++) {
        query[Object.keys(data)[i]] = data[Object.keys(data)[i]];
    }
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata = await dbo.collection(tableName).find(query).toArray();
        if (resultdata != undefined && resultdata.length > 0) {
            var totalRow = await dbo.collection(tableName).countDocuments(query);
            returnObject.TotalRows = totalRow;
            returnObject.data = resultdata;
        }
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.DeleteData = async function(data, tableName) {
    var returnObject = { error: 0, msg: '', data: null };
    var query = {};
    for (var i = 0; i < Object.keys(data).length; i++) {
        query[Object.keys(data)[i]] = data[Object.keys(data)[i]];
    }
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        await dbo.collection(tableName).deleteMany(query);
    } catch (err) {
        console.error('mongoData.DeleteData' + err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.ReadCache = async function(keycache) {
    var returnObject = { error: 0, msg: '', data: null };
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata = await dbo.collection(CacheTableName).find({ _id: keycache }).toArray();
        if (resultdata != undefined && resultdata.length > 0) {
            returnObject.data = resultdata;
        }
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};

mongoData.ConnectData = async function() {
    var db;
    try {
        db = await mongoClient.connect(url, options);
    } catch (err) {

    } finally {
        return db;
    }
};

mongoData.InsertData = async function(db, collection, objectdata) {
    var resultInsert = 0;
    try {
        var dbo = db.db(dbName);
        await dbo.collection(collection).insertOne(objectdata);
        resultInsert = 1;
    } catch (err) {
        console.error(err);
    } finally {
        return resultInsert;
    }
};

mongoData.Insert = async function(objectdata, tableName) {
    var returnObject = { error: 0, msg: '', data: null };
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        await dbo.collection(tableName).insertOne(objectdata);
        //console.log('insert done _id ' + objectdata._id);
        returnObject.data = objectdata;
        db.close();
    } catch (err) {
        console.error(err);
        returnObject.msg = err;
        returnObject.error = 1;
    } finally {
        if (db) {
            try {
                await db.close();
                return returnObject;
            } catch (err) {
                console.error(err);
                returnObject.msg = err;
                returnObject.error = 1;
                return returnObject;
            }
        } else {
            returnObject.msg = 'can not connect db mongo';
            returnObject.error = 1;
            return returnObject;
        }
    }
};