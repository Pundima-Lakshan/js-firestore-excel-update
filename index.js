import XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import {
  getFirestore,
  doc,
  setDoc,
  updateDoc,
  collection,
  getDocs,
} from "@firebase/firestore";

// Initialize Firebase
const firebaseConfig = {
  apiKey: "",
  authDomain: "",
  projectId: "",
  storageBucket: "",
  messagingSenderId: "",
  appId: "",
};
const app = initializeApp(firebaseConfig);

const letter = "E";

// Read data from Excel file
const workbook = XLSX.readFile(`${letter}.xlsx`);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Create collection with same name as Excel file and write data to Firestore
const collectionName = "demoGuests";
const db = getFirestore(app);
const collectionRef = collection(db, collectionName);

// Check if collection exists
getDocs(collectionRef)
  .then((querySnapshot) => {
    if (querySnapshot.size === 0) {
      // Collection does not exist, create new documents
      data.slice(1).forEach((row, index) => {
        let id = `${letter}` + (index + 1).toString().padStart(3, "0");
        const docRef = doc(collectionRef, id);
        const dataObj = {};
        row.forEach((val, i) => {
          const field = data[0][i];
          dataObj[field] = val;
        });
        setDoc(docRef, dataObj);
      });
    } else {
      // Collection exists, add new documents or update existing ones
      data.slice(1).forEach((row, index) => {
        let id = `${letter}` + (index + 1).toString().padStart(3, "0");
        const docRef = doc(collectionRef, id);
        const dataObj = {};
        row.forEach((val, i) => {
          const field = data[0][i];
          dataObj[field] = val;
        });
        const docExists = querySnapshot.docs.some((doc) => doc.id === id);
        if (docExists) {
          // Document already exists, merge fields
          updateDoc(docRef, dataObj);
        } else {
          // Document does not exist, create new document
          setDoc(docRef, dataObj);
        }
      });
    }
  })
  .catch((error) => {
    console.log("Error getting collection documents: ", error);
  });
