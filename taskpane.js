// ===== FIREBASE CONFIG (YOUR PROJECT) =====
const firebaseConfig = {
  apiKey: "AIzaSyArrOnmin5gtNclRVVGc1qd__f8RNr_QjA",
  authDomain: "outlook-connector-cb347.firebaseapp.com",
  projectId: "outlook-connector-cb347",
  storageBucket: "outlook-connector-cb347.firebasestorage.app",
  messagingSenderId: "1024064864900",
  appId: "1:1024064864900:web:61f34f3296251c7e92f8d1"
};

// INIT FIREBASE
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// GET EMAIL ID
function getEmailId() {
  return Office.context.mailbox.item.itemId;
}

// LOAD NOTE
async function loadNote() {
  const id = getEmailId();

  const doc = await db.collection("notes").doc(id).get();

  if (doc.exists) {
    document.getElementById("noteBox").value = doc.data().text;
  }
}

// SAVE NOTE
async function saveNote() {
  const id = getEmailId();
  const text = document.getElementById("noteBox").value;

  await db.collection("notes").doc(id).set({
    text: text,
    updated: new Date().toISOString()
  });

  alert("Saved!");
}

// START
Office.onReady(() => {
  loadNote();
});