// Pak QR Print - Firebase config (TEMPLATE)
// ============================================================
// IS FILE KO PEHLE BHARNA LAZMI HAI - warna site kaam nahi karegi.
// Firebase Console (console.firebase.google.com) se apne project ki
// config yahan paste karein. Har jagah "YAHAN-..." likha hai - us ko
// apni asal value se badlein. Ye file GitHub par PUBLIC hoti hai -
// is mein sirf yehi config hoti hai, koi password nahi hota.
// Dukaan ka password Firebase Auth mein banta hai (Hidayat dekhein),
// yahan nahi likha jata.
// ============================================================
window.PAKQR_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBZXoguVL0KT_9V0xUwYM7DqqBEM4NZ9sQ",
  authDomain: "pak-qr-print.firebaseapp.com",
  databaseURL: "https://pak-qr-print-default-rtdb.firebaseio.com",
  projectId: "pak-qr-print",
  storageBucket: "pak-qr-print.firebasestorage.app",
  messagingSenderId: "88024052278",
  appId: "1:88024052278:web:e029128a284d88cc649e97"
};

// Dukaan ka ID - chhota, English, baghair space (masalan: pakqrprint).
// Customer link isi se banta hai: index.html?shop=pakqrprint
window.PAKQR_SHOP_ID = "pakqrprint";

// Customer page ka link (QR poster isi ka banta hai).
// Khaali chhodein to site khud apna address use kar legi.
// Apna GitHub Pages link likhna ho to yahan likhein, masalan:
// "https://USERNAME.github.io/pak-qr-print"
window.PAKQR_CUSTOMER_URL = "";
