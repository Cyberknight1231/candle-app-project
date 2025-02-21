import { initializeApp } from 'firebase/app';
import { getFirestore, Firestore } from 'firebase/firestore';

// Your Firebase configuration object
const firebaseConfig = {
  apiKey: "AIzaSyAx-NQvh6k1nMJTSlqM4YZ93_C16gx1Qqk",
  authDomain: "candle-project-bd8d8.firebaseapp.com",
  projectId: "candle-project-bd8d8",
  storageBucket: "candle-project-bd8d8.appspot.com",
  messagingSenderId: "687448548450",
  appId: "1:687448548450:web:f5e1d4e97a7e63edc1dd3b"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore
const db: Firestore = getFirestore(app);

// Export Firestore instance
export { db };