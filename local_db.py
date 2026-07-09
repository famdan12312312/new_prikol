import sqlite3
import json
import uuid
import os
import datetime
class ObjectId:
    def __init__(self, oid=None):
        if oid is None:
            # Generate a 24-character hex string resembling ObjectId
            import secrets
            self._str = secrets.token_hex(12)
        else:
            if isinstance(oid, ObjectId):
                self._str = oid._str
            elif isinstance(oid, str):
                if len(oid) == 24 and all(c in '0123456789abcdefABCDEF' for c in oid):
                    self._str = oid.lower()
                else:
                    raise ValueError(f"'{oid}' is not a valid ObjectId")
            else:
                raise TypeError(f"id must be a string or ObjectId, not {type(oid)}")

    def __str__(self):
        return self._str

    def __repr__(self):
        return f"ObjectId('{self._str}')"

    def __eq__(self, other):
        if isinstance(other, ObjectId):
            return self._str == other._str
        if isinstance(other, str):
            return self._str == other
        return False

    def __hash__(self):
        return hash(self._str)

class Cursor:
    def __init__(self, items):
        self.items = items

    def __iter__(self):
        return iter(self.items)
    
    def __list__(self):
        return list(self.items)

def match_query(doc, query):
    if not query:
        return True
    
    for key, condition in query.items():
        val = doc.get(key)
        if isinstance(condition, dict):
            # Handle operators like $gt, $in, etc.
            if "$gt" in condition:
                if val is None or not (val > condition["$gt"]):
                    return False
            if "$lt" in condition:
                if val is None or not (val < condition["$lt"]):
                    return False
            if "$in" in condition:
                in_list = [str(x) if isinstance(x, ObjectId) else x for x in condition["$in"]]
                val_str = str(val) if isinstance(val, ObjectId) else val
                if val_str not in in_list:
                    return False
            if "$exists" in condition:
                exists = key in doc
                if exists != condition["$exists"]:
                    return False
        else:
            if isinstance(val, ObjectId) or isinstance(condition, ObjectId):
                if str(val) != str(condition):
                    return False
            else:
                if val != condition:
                    return False
    return True

class Collection:
    def __init__(self, db_name, name, conn):
        self.db_name = db_name
        self.name = name
        self.conn = conn
        self._ensure_table()

    def _ensure_table(self):
        cursor = self.conn.cursor()
        # Create a table for each collection. We will store data as JSON.
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {self.name} (
                id TEXT PRIMARY KEY,
                data TEXT
            )
        """)
        self.conn.commit()

    def insert_one(self, document):
        if "_id" not in document:
            document["_id"] = ObjectId()
        
        doc_id = str(document["_id"])
        
        # Serialize to JSON, converting ObjectId to string for storage
        def default_encoder(o):
            if isinstance(o, ObjectId):
                return str(o)
            if isinstance(o, (datetime.datetime, datetime.date)):
                return o.isoformat()
            raise TypeError(f"Object of type {o.__class__.__name__} is not JSON serializable")
            
        data_json = json.dumps(document, default=default_encoder)
        
        cursor = self.conn.cursor()
        cursor.execute(f"INSERT OR REPLACE INTO {self.name} (id, data) VALUES (?, ?)", (doc_id, data_json))
        self.conn.commit()
        
        class InsertOneResult:
            inserted_id = document["_id"]
        return InsertOneResult()

    def insert_many(self, documents):
        inserted_ids = []
        for doc in documents:
            self.insert_one(doc)
            inserted_ids.append(doc["_id"])
        
        class InsertManyResult:
            def __init__(self, ids):
                self.inserted_ids = ids
        return InsertManyResult(inserted_ids)

    def _get_all_docs(self):
        cursor = self.conn.cursor()
        cursor.execute(f"SELECT data FROM {self.name}")
        rows = cursor.fetchall()
        docs = []
        for row in rows:
            doc = json.loads(row[0])
            # Convert string _id back to ObjectId if possible
            if "_id" in doc:
                try:
                    doc["_id"] = ObjectId(doc["_id"])
                except Exception:
                    pass
            docs.append(doc)
        return docs

    def find(self, query=None, projection=None):
        docs = self._get_all_docs()
        filtered = [doc for doc in docs if match_query(doc, query)]
        
        # Apply projection if specified
        if projection:
            projected_docs = []
            for doc in filtered:
                new_doc = {}
                include_mode = any(v for k, v in projection.items() if k != "_id" and v)
                
                if include_mode:
                    for k, v in projection.items():
                        if v and k in doc:
                            new_doc[k] = doc[k]
                    if projection.get("_id", 1) and "_id" in doc:
                        new_doc["_id"] = doc["_id"]
                else:
                    new_doc = doc.copy()
                    for k, v in projection.items():
                        if not v and k in new_doc:
                            del new_doc[k]
                projected_docs.append(new_doc)
            return Cursor(projected_docs)

        return Cursor(filtered)

    def find_one(self, query=None):
        docs = self._get_all_docs()
        for doc in docs:
            if match_query(doc, query):
                return doc
        return None

    def delete_many(self, query=None):
        if not query:
            # Delete all
            cursor = self.conn.cursor()
            cursor.execute(f"DELETE FROM {self.name}")
            self.conn.commit()
            
            class DeleteResult:
                deleted_count = cursor.rowcount
            return DeleteResult()
            
        docs = self._get_all_docs()
        deleted_count = 0
        cursor = self.conn.cursor()
        
        for doc in docs:
            if match_query(doc, query):
                doc_id = str(doc.get("_id"))
                cursor.execute(f"DELETE FROM {self.name} WHERE id = ?", (doc_id,))
                deleted_count += 1
                
        self.conn.commit()
        class DeleteResult:
            def __init__(self, count):
                self.deleted_count = count
        return DeleteResult(deleted_count)

    def update_one(self, query, update):
        docs = self._get_all_docs()
        for doc in docs:
            if match_query(doc, query):
                # Apply update
                if "$set" in update:
                    for k, v in update["$set"].items():
                        doc[k] = v
                if "$addToSet" in update:
                    for k, v in update["$addToSet"].items():
                        if k not in doc:
                            doc[k] = []
                        if isinstance(doc[k], list) and v not in doc[k]:
                            doc[k].append(v)
                
                # Save back to db
                doc_id = str(doc["_id"])
                
                def default_encoder(o):
                    if isinstance(o, ObjectId):
                        return str(o)
                    if isinstance(o, (datetime.datetime, datetime.date)):
                        return o.isoformat()
                    raise TypeError(f"Object of type {o.__class__.__name__} is not JSON serializable")
                    
                data_json = json.dumps(doc, default=default_encoder)
                cursor = self.conn.cursor()
                cursor.execute(f"UPDATE {self.name} SET data = ? WHERE id = ?", (data_json, doc_id))
                self.conn.commit()
                
                class UpdateResult:
                    matched_count = 1
                    modified_count = 1
                return UpdateResult()
                
        class UpdateResult:
            matched_count = 0
            modified_count = 0
        return UpdateResult()
        
    def drop(self):
        cursor = self.conn.cursor()
        cursor.execute(f"DROP TABLE IF EXISTS {self.name}")
        self.conn.commit()
        self._ensure_table()

class Database:
    def __init__(self, name, conn):
        self.name = name
        self.conn = conn
        self.collections = {}

    def __getattr__(self, name):
        return self[name]

    def __getitem__(self, name):
        if name not in self.collections:
            self.collections[name] = Collection(self.name, name, self.conn)
        return self.collections[name]

class MongoClient:
    def __init__(self, host=None, port=None, db_file="university_db.sqlite", **kwargs):
        self.db_file = db_file
        self.conn = sqlite3.connect(self.db_file, check_same_thread=False)
        self.databases = {}

    def __getattr__(self, name):
        return self[name]

    def __getitem__(self, name):
        if name not in self.databases:
            self.databases[name] = Database(name, self.conn)
        return self.databases[name]
    
    def server_info(self):
        return {"version": "sqlite-local-proxy"}

    def close(self):
        self.conn.close()
