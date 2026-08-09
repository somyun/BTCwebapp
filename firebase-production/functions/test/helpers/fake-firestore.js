"use strict";

function clone(value) {
  return value === undefined ? undefined : structuredClone(value);
}

class FakeSnapshot {
  constructor(reference, data) {
    this.ref = reference;
    this.id = reference.id;
    this.exists = data !== undefined;
    this._data = clone(data);
  }

  data() {
    return clone(this._data);
  }

  get(field) {
    return this._data?.[field];
  }
}

class FakeDocumentReference {
  constructor(firestore, path) {
    this.firestore = firestore;
    this.path = path;
    this.id = path.split("/").at(-1);
  }

  async get() {
    return new FakeSnapshot(this, this.firestore.documents.get(this.path));
  }

  async set(data) {
    this.firestore.documents.set(this.path, clone(data));
  }

  async delete() {
    this.firestore.documents.delete(this.path);
  }

  collection(name) {
    return new FakeCollectionReference(this.firestore, `${this.path}/${name}`);
  }
}

class FakeCollectionReference {
  constructor(firestore, path) {
    this.firestore = firestore;
    this.path = path;
  }

  doc(id) {
    return new FakeDocumentReference(this.firestore, `${this.path}/${id}`);
  }

  async get() {
    const prefix = `${this.path}/`;
    const docs = [];
    for (const [path, data] of this.firestore.documents.entries()) {
      const suffix = path.startsWith(prefix) ? path.slice(prefix.length) : "";
      if (suffix && !suffix.includes("/")) {
        docs.push(new FakeSnapshot(new FakeDocumentReference(this.firestore, path), data));
      }
    }
    return { docs };
  }
}

class FakeFirestore {
  constructor(seed = {}) {
    this.documents = new Map(Object.entries(seed).map(([path, data]) => [path, clone(data)]));
  }

  collection(name) {
    return new FakeCollectionReference(this, name);
  }
}

module.exports = { FakeFirestore };
