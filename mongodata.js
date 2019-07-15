var mongoClient = require('mongodb').MongoClient;
var url = "mongodb://127.0.0.1:27017/";
var dbName = "crmCache";
var options = {
    useNewUrlParser: true
};
var CacheTableName = "CacheData";
var SyncTableName = "Customer";
var mongoData = (exports.mongoData = {});
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

mongoData.BuildQuery = async function(objectCondition) {
    var query = {};
    console.log(objectCondition);
    for (var i = 0; i < objectCondition.length; i++) {
        switch (objectCondition[i].ConditionName.toUpperCase()) {
            case "ID_NO":
                objectCondition[i].ConditionName = "IdNo";
                break;
            case "PRIORITY_TYPE":
                objectCondition[i].ConditionName = "PriorityType";
                break;
            case "FULL_NAME":
                objectCondition[i].ConditionName = "FullName";
                break;
            case "CIF_NO":
                objectCondition[i].ConditionName = "CifNo";
                break;
            case "BUSINESS_REGIST_CER_NO":
                objectCondition[i].ConditionName = "BusinessRegistCerNo";
                break;

        }
        if (objectCondition[i].ConditionOp.toUpperCase() == 'LIKE' || objectCondition[i].ConditionOp.toUpperCase() == 'BETWEEN_NUMBERS') {
            switch (objectCondition[i].ConditionOp.toUpperCase()) {
                case "LIKE":
                    switch (objectCondition[i].ConditionName.toUpperCase()) {
                        case "PHONE_NO":
                            break;
                            // case "ID_NO":
                            // query["IdNo"]= { $regex: objectCondition[i].ConditionValue1, $options: 'i' };
                            // break;
                        default:
                            query[objectCondition[i].ConditionName] = { $regex: objectCondition[i].ConditionValue1, $options: 'i' };
                    }
                    break;
                case "BETWEEN_NUMBERS":
                    query[objectCondition[i].ConditionName] = { $gt: objectCondition[i].ConditionValue1, $lt: objectCondition[i].ConditionValue2 };
                    break;
            }
        } else {
            switch (objectCondition[i].ConditionName.toUpperCase()) {
                // case "PRIORITY_TYPE":
                // query["PriorityType"]= objectCondition[i].ConditionValue1;
                // break;
                case "BIRTH_DAY":
                    query[objectCondition[i].ConditionName] = { $gt: new Date(objectCondition[i].ConditionValue1), $lt: new Date(objectCondition[i].ConditionValue2) };
                    break;
                case "PHONE_NO":
                    query[objectCondition[i].ConditionName] = { $regex: objectCondition[i].ConditionValue1, $options: 'i' };
                    break;
                case "OPP_STATUS":
                    //var dbSQL = "select case when exists (SELECT 1  FROM CRM_CUST_OPPOTUNITIES o WHERE c.CRM_CUST_ID=o.CRM_CUST_ID and  o.STATUS_ID = " + objectCondition[i].ConditionValue1 + ") then 1 else 0 from DUAL";
                    query["OPPOTUNITIES.STATUS_ID"] = objectCondition[i].ConditionValue1;
                    break;
                case "ASSIGN_RM_STATUS":
                    // if i.ConditionValue1 =0 then
                    // v_strWhere := v_strWhere || ' AND (  c.ACCOUNT_OFFICER  =''' ||  i.ConditionValue2 || ''' or c.ASSIGN_RM_STATUS =0) '   ;
                    // else
                    // v_strWhere := v_strWhere || ' AND c.ACCOUNT_OFFICER IS NOT NULL';
                    // end if;
                    if (objectCondition[i].ConditionValue1 == 0) {
                        query["ASSIGN_RM_STATUS"] = [{ "AccountOfficer": objectCondition[i].ConditionValue1 }, { "AssignRmStatus": 0 }];
                    } else {
                        query["AssignRmStatus"] = { $ne: null };
                    }

                    break;
                case "OFFICER_NO":
                    // v_strWhere := v_strWhere || ' AND (c.OFFICER_NO = ''' || i.ConditionValue1 || ''' or
                    // iSExistOppotunitiesCust(nvl(C.CRM_CUST_ID,CIF_NO), '''|| i.ConditionValue1  || ''') >0)' || UTL_tcp.CRLF;
                    query["OFFICER_NO"] = [{ "AccountOfficer": objectCondition[i].ConditionValue1 }, { $where: "this.OPPOTUNITIES.length > 0" }];
                    break;
                case "STATUS_ID":
                    // if instr( i.ConditionValue1,';')>0 then
                    // v_strWhere := v_strWhere || ' AND  instr(''' ||  i.ConditionValue1 || ''',  c.STATUS_ID  )>0 ' ;
                    // else
                    // v_strWhere := v_strWhere || ' AND  c.STATUS_ID =''' ||  i.ConditionValue1 || '''';
                    if (objectCondition[i].ConditionValue1.indexOf(';') >= 0) {
                        query[objectCondition[i].ConditionName] = { $regex: objectCondition[i].ConditionValue1, $options: 'i' };
                    } else {
                        query[objectCondition[i].ConditionName] = objectCondition[i].ConditionValue1;
                    }

                    break;
                case "CONTACT_STATUS_ID":
                    query["ContactStatusId"] = objectCondition[i].ConditionValue1;
                    // v_strWhere := v_strWhere || ' AND c.CONTACT_STATUS_ID = ' || i.ConditionValue1;
                    break;
                default:
                    switch (objectCondition[i].ConditionOp.toUpperCase()) {
                        case "=":
                            query[objectCondition[i].ConditionName] = objectCondition[i].ConditionValue1;
                            break;
                        case "<>":
                            query[objectCondition[i].ConditionName] = { $ne: objectCondition[i].ConditionValue1 };
                            break;
                        case "!=":
                            query[objectCondition[i].ConditionName] = { $ne: objectCondition[i].ConditionValue1 };
                            break;
                        case "<":
                            query[objectCondition[i].ConditionName] = { $lt: objectCondition[i].ConditionValue1 };
                            break;
                        case ">":
                            query[objectCondition[i].ConditionName] = { $gt: objectCondition[i].ConditionValue1 };
                            break;
                        default:
                            query[objectCondition[i].ConditionName] = objectCondition[i].ConditionValue1;
                    }
            }
        }
    }
    return query;
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

mongoData.QueryData = async function(query, collection, pageindex, pagesize) {
    var returnObject = { error: 0, msg: '', TotalRows: 0, data: null };
    var skipdata = (parseInt(pageindex) - 1) * parseInt(pagesize);
    var db = await mongoClient.connect(url, options);
    try {
        var dbo = db.db(dbName);
        var resultdata = await dbo.collection(collection).find(query).skip(skipdata).limit(parseInt(pagesize)).toArray();
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

mongoData.DeleteData = async function(db, recordtablename, recordid) {
    var returnObject = { error: 0, msg: '', data: null };
    //var db = await mongoClient.connect(url, options);
    try {
        var nameElement = "CrmCustId";
        if (recordtablename == "CustomerPros") {
            nameElement = "CifNo";
        }
        var dbo = db.db(dbName);
        await dbo.collection(recordtablename).deleteOne({
            nameElement: recordid
        });
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