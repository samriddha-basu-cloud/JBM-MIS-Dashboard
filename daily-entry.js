import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js';
import { getFirestore, collection, addDoc, getDocs, updateDoc, deleteDoc, doc, query, where, orderBy, Timestamp, setDoc } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js';

const firebaseConfig = {
    apiKey: "AIzaSyDTczWACmsAcYKY743wLqNe0DiRAy5k6tk",
    authDomain: "mis-dashboard-f1688.firebaseapp.com",
    projectId: "mis-dashboard-f1688",
    storageBucket: "mis-dashboard-f1688.firebasestorage.app",
    messagingSenderId: "476511626207",
    appId: "1:476511626207:web:d3eb5f41cbec14b7325a6a"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const WORKING_MINUTES = 1335;
const EFFICIENCY = 0.85;

const MASTER_DATA = [{"srNo":1,"assyCategory":"BIW","sapCode":"1044911","partNo":"0103AS201010N","partDescription":"0103AS201010N-BEAM ASSY FRT DR INTRSN RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"MANUAL","stdManhead":"3","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":2,"assyCategory":"BIW","sapCode":"1083474","partNo":"0101CS201830N","partDescription":"FG-S210 RENF ASY FRT TUNEL FRT CS201830N","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"NEW SHED","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP10","stdManhead":"8","cycleTime":"2.7","jph":"19.1","capacityPerDay":"425.0"},{"srNo":3,"assyCategory":"BIW","sapCode":"1044861","partNo":"0101BS201940N","partDescription":"0101BS201940N-REINF ASSY-SIDE SILL-LH.1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (New)","operation":"OP10","stdManhead":"12","cycleTime":"2.8","jph":"18.3","capacityPerDay":"406.3"},{"srNo":4,"assyCategory":"BIW","sapCode":"1044862","partNo":"0101BS201930N","partDescription":"0101BS201930N-REINF ASSY-SIDE SILL-RH.1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"NEW SHED","ownership":"JBM","processType":"Robotic","cellNo":"Cell 4 (New)","operation":"OP10","stdManhead":"12","cycleTime":"2.8","jph":"18.0","capacityPerDay":"401.2"},{"srNo":5,"assyCategory":"BIW","sapCode":"1044866","partNo":"0101BS202220N","partDescription":"0101BS202220N-PNL ASSY-SIDE SILL INR-RH1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (New)","operation":"OP10","stdManhead":"4","cycleTime":"2.8","jph":"18.3","capacityPerDay":"407.2"},{"srNo":6,"assyCategory":"BIW","sapCode":"1044909","partNo":"0103AS200990N","partDescription":"0103AS200990N-BEAM ASSY FRT DR INTRSN RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"MIG WELDING","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (New)","operation":"MANUAL","stdManhead":"12","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":7,"assyCategory":"BIW","sapCode":"1044781","partNo":"0102AS200230N","partDescription":"0102AS200230N-MBR ASSY FENDER APRON LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (New)","operation":"MANUAL","stdManhead":"16","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":8,"assyCategory":"BIW","sapCode":"1044863","partNo":"0101DS200380N","partDescription":"0101DS200380N-RAIL ASSY-ROOF RR","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"8","cycleTime":"2.8","jph":"18.3","capacityPerDay":"406.3"},{"srNo":9,"assyCategory":"BIW","sapCode":"1044910","partNo":"0103AS201000N","partDescription":"0103AS201000N-BEAM ASSY RR DR INTRSN LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"MIG WELDING","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"MANUAL","stdManhead":"12","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":10,"assyCategory":"BIW","sapCode":"1044884","partNo":"0101ES201120N","partDescription":"0101ES201120N-PNL ASSY WHL HSE RR INR LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 10","operation":"OP 10/ OP 20","stdManhead":"8","cycleTime":"2.8","jph":"18.1","capacityPerDay":"402.9"},{"srNo":11,"assyCategory":"BIW","sapCode":"1044908","partNo":"0103AS200980N","partDescription":"0103AS200980N-BEAM ASSY FRT DR INTRSN LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"MIG WELDING","ownership":"JBM","processType":"Manual","cellNo":"Cell 10","operation":"MANUAL","stdManhead":"12","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":12,"assyCategory":"BIW","sapCode":"1044869","partNo":"0101BS202040N","partDescription":"0101BS202040N-MBR ASSY FRT SIDE OTR LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 10","operation":"MANUAL","stdManhead":"16","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":13,"assyCategory":"BIW","sapCode":"1044885","partNo":"0101ES201110N","partDescription":"0101ES201110N-PNL ASSY WHL HSE RR INR RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 11","operation":"OP 10/ OP 20","stdManhead":"16","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":14,"assyCategory":"BIW","sapCode":"1044886","partNo":"0101ES201010N","partDescription":"0101ES201010N-REINF ASSY-B PLR OTR-LH1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"20","cycleTime":"2.4","jph":"21.2","capacityPerDay":"472.6"},{"srNo":15,"assyCategory":"BIW","sapCode":"1089403","partNo":"0101DS200710N","partDescription":"REINF ASSY SUNROOF 0101DS200710N","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"MANUAL","stdManhead":"16","cycleTime":"3.8","jph":"13.4","capacityPerDay":"297.5"},{"srNo":16,"assyCategory":"BIW","sapCode":"1044780","partNo":"0102AS200240N","partDescription":"0102AS200240N-MBR ASSY FENDER APRON RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"MANUAL","stdManhead":"16","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":17,"assyCategory":"BIW","sapCode":"1044887","partNo":"0101ES201020N","partDescription":"0101ES201020N-REINF ASSY-B PLR OTR-RH1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"MANUAL","stdManhead":"20","cycleTime":"2.4","jph":"21.2","capacityPerDay":"472.6"},{"srNo":18,"assyCategory":"BIW","sapCode":"1044868","partNo":"0101ES201310N","partDescription":"0101ES201310N-PNL COMP COWL SIDE INR LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 10","stdManhead":"16","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":19,"assyCategory":"BIW","sapCode":"1044890","partNo":"0101CS200560N","partDescription":"0101CS200560N-MBR ASSY HOOD HINGE MTG LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 10","stdManhead":"8","cycleTime":"2.6","jph":"19.7","capacityPerDay":"438.6"},{"srNo":20,"assyCategory":"BIW","sapCode":"1044891","partNo":"0101CS200570N","partDescription":"0101CS200570N-MBR ASSY HOOD HINGE MTG RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 10","stdManhead":"3","cycleTime":"2.6","jph":"19.7","capacityPerDay":"438.6"},{"srNo":21,"assyCategory":"BIW","sapCode":"1044867","partNo":"0101ES201320N","partDescription":"0101ES201320N-PNL COMP COWL SIDE INR LH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"MANUAL","operation":"OP 10","stdManhead":"4","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":22,"assyCategory":"BIW","sapCode":"1044865","partNo":"0101BS202210N","partDescription":"0101BS202210N-PNL ASSY-SIDE SILL INR-LH1","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"3","cycleTime":"2.8","jph":"18.0","capacityPerDay":"401.2"},{"srNo":23,"assyCategory":"BIW","sapCode":"1044870","partNo":"0101BS202050N","partDescription":"0101BS202050N-MBR ASSY FRT SIDE OTR RH","model":"XUV 3XO + XUV 4OO","variant":"3XO","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"54","cycleTime":"3.0","jph":"17.2","capacityPerDay":"382.5"},{"srNo":24,"assyCategory":"BIW","sapCode":"1044778","partNo":"0101DS200200N","partDescription":"0101DS200200N-RAIL ASSY ROOF NO2 CTR","model":"OPTION","variant":"3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 10","stdManhead":"4","cycleTime":"6.7","jph":"7.6","capacityPerDay":"170.0"},{"srNo":25,"assyCategory":"BIW","sapCode":"1068639","partNo":"0101BS205170N","partDescription":"FG-S210 ASY MB DASH LWR INR LH","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 10","stdManhead":"3","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":26,"assyCategory":"BIW","sapCode":"1068640","partNo":"0101BS205180N","partDescription":"FG-S210 ASY MB DASH LWR INR RH","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 10","stdManhead":"2","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":27,"assyCategory":"BIW","sapCode":"1068644","partNo":"0101BS206160N","partDescription":"FG-S210 BKT ASY BATT RR RH","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":28,"assyCategory":"BIW","sapCode":"1068643","partNo":"0101BS206170N","partDescription":"FG-S210 BKT ASY BATT RR LH","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"1.5","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":29,"assyCategory":"BIW","sapCode":"1044912","partNo":"0101BS202380N","partDescription":"MBR ASSLY RR SIDE LH (REG)","model":"XUV 3XO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":30,"assyCategory":"BIW","sapCode":"1078284","partNo":"0101BS206800N","partDescription":"MBR ASSLY RR SIDE LH (EXPORT)","model":"EXPORT","variant":"EXPORT","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":31,"assyCategory":"BIW","sapCode":"1068642","partNo":"0101BS204500N","partDescription":"MBR ASSLY RR SIDE LH (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":32,"assyCategory":"BIW","sapCode":"1044913","partNo":"0101BS202390N","partDescription":"MBR ASSLY RR SIDE RH (REG)","model":"XUV 3XO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":33,"assyCategory":"BIW","sapCode":"1078283","partNo":"0101BS206810N","partDescription":"MBR ASSLY RR SIDE RH (EXPORT)","model":"EXPORT","variant":"EXPORT","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":34,"assyCategory":"BIW","sapCode":"1068641","partNo":"0101BS204520N","partDescription":"MBR ASSLY RR SIDE RH (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":35,"assyCategory":"BIW","sapCode":"1044888","partNo":"0101BS202060N","partDescription":"MBR ASSLY FRT SIDE INR LH (REG)","model":"XUV 3XO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":36,"assyCategory":"BIW","sapCode":"1068648","partNo":"0101BS205070N","partDescription":"MBR ASSLY FRT SIDE INR LH (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":37,"assyCategory":"BIW","sapCode":"1044889","partNo":"0101BS202070N","partDescription":"MBR ASSLY FRT SIDE INR RH (REG)","model":"XUV 3XO","variant":"3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":38,"assyCategory":"BIW","sapCode":"1068649","partNo":"0101BS205080N","partDescription":"MBR ASSLY FRT SIDE INR RH (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":39,"assyCategory":"BIW","sapCode":"1044883","partNo":"0101CS200540N","partDescription":"PNL ASSY COWL LWR ","model":"XUV 3XO + XUV 4OO","variant":"XUV 400","prodShopArea":"XUV 3XO + XUV 4OO","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":40,"assyCategory":"BIW","sapCode":"1062277","partNo":"0101CS201160N","partDescription":"PNL ASSY COWL LWR (LHD)","model":"LHD","variant":"LHD","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":41,"assyCategory":"BIW","sapCode":"1096677","partNo":"0101BS207090N","partDescription":"REINF ASSY FRT TUNNEL (REG)","model":"XUV 3XO","variant":"XUV 3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":42,"assyCategory":"BIW","sapCode":"1068647","partNo":"0101BS204080N","partDescription":"REINF ASSY FRT TUNNEL (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 20","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":43,"assyCategory":"BIW","sapCode":"1044882","partNo":"0101CS200530N","partDescription":"MBR ASSY DASH LWR FRT (REG)","model":"XUV 3XO","variant":"XUV 3XO","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":44,"assyCategory":"BIW","sapCode":"1055042","partNo":"0102CS200160N","partDescription":"MBR ASSY DASH LWR FRT (LHD)","model":"LHD","variant":"LHD","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":45,"assyCategory":"BIW","sapCode":"1044786","partNo":"0101BS202150N","partDescription":"MEMBER ASSY CTR SIDE","model":"XUV 3XO","variant":"XUV 3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":46,"assyCategory":"BIW","sapCode":"1068646","partNo":"0101BS204440N","partDescription":"MEMBER ASSY CTR SIDE (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":47,"assyCategory":"BIW","sapCode":"1044864","partNo":"0101BS202240N","partDescription":"CROSS MBR ASSY-RR FLR NO2","model":"XUV 3XO","variant":"XUV 3XO","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":48,"assyCategory":"BIW","sapCode":"1068645","partNo":"0101BS204230N","partDescription":"CROSS MBR ASSY-RR FLR NO2 (EV)","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 10","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":49,"assyCategory":"BIW","sapCode":"1044783","partNo":"0101CS200590N","partDescription":"REINF ASSY DASH","model":"XUV 3XO","variant":"XUV 3XO","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":50,"assyCategory":"BIW","sapCode":"1055044","partNo":"0101CS200630N","partDescription":"REINF ASSY DASH (LHD)","model":"LHD","variant":"LHD","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.8","jph":"18.1","capacityPerDay":"403.8"},{"srNo":51,"assyCategory":"BIW","sapCode":"1096218","partNo":"0102FW500030N","partDescription":"0102FW500030NASSY HEAD LAMP RH","model":"W502","variant":"W502","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"3","cycleTime":"3.4","jph":"15.0","capacityPerDay":"334.1"},{"srNo":52,"assyCategory":"BIW","sapCode":"1096224","partNo":"0101BW502700N","partDescription":"0101BW502700N PNLASY LONG MBR RR EXTN LH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"3.4","jph":"15.0","capacityPerDay":"334.1"},{"srNo":53,"assyCategory":"BIW","sapCode":"1096223","partNo":"0101BW502750N","partDescription":"0101BW502750N PNLASY LONG MBR RR EXTN RH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"5","cycleTime":"3.4","jph":"15.0","capacityPerDay":"334.1"},{"srNo":54,"assyCategory":"BIW","sapCode":"1096222","partNo":"0101BW502880N","partDescription":"0101BW502880NCROSS MBR ASSY REAR SEAT","model":"W502","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"3","cycleTime":"3.7","jph":"13.8","capacityPerDay":"306.9"},{"srNo":55,"assyCategory":"BIW","sapCode":"1096219","partNo":"0102FW500040N","partDescription":"0102FW500040NASSY HEAD LAMP LH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"2.6","jph":"19.5","capacityPerDay":"434.4"},{"srNo":56,"assyCategory":"BIW","sapCode":"1096227","partNo":"0101BW504420N","partDescription":"0101BW504420NBRACKET ASSY PARKING BRAKE","model":"W502","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"3","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":57,"assyCategory":"BIW","sapCode":"1096217","partNo":"0101EW503580N","partDescription":"0101EW503580N REINF ASSY PLR A INR LH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4.3","cycleTime":"3.7","jph":"13.8","capacityPerDay":"306.9"},{"srNo":58,"assyCategory":"BIW","sapCode":"1096225","partNo":"0101EW503100N","partDescription":"0101EW503100NPANEL ASSY CANTRAIL RH","model":"W502","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"5","cycleTime":"3.2","jph":"15.7","capacityPerDay":"350.2"},{"srNo":59,"assyCategory":"BIW","sapCode":"1096221","partNo":"0101BW504330N","partDescription":"0101BW504330NCROSS MBR ASSY SKID REAR","model":"W502","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"3.7","jph":"13.8","capacityPerDay":"306.9"},{"srNo":60,"assyCategory":"BIW","sapCode":"1096226","partNo":"0101EW503110N","partDescription":"0101EW503110NPANEL ASSY CANTRAIL LH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"1","cycleTime":"3.7","jph":"13.8","capacityPerDay":"306.9"},{"srNo":61,"assyCategory":"BIW","sapCode":"1096216","partNo":"0101EW503570N","partDescription":"0101EW503570N REINF ASSY PLR A INR RH","model":"W502","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"6","cycleTime":"3.2","jph":"15.7","capacityPerDay":"350.2"},{"srNo":62,"assyCategory":"TCF","sapCode":"1043879","partNo":"0119BAL00350N","partDescription":"FG-ASSY RR BMPR SUPT STRCR-0119BAL00350N","model":"SC + DC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"5","cycleTime":"16.7","jph":"3.1","capacityPerDay":"68.0"},{"srNo":63,"assyCategory":"BIW","sapCode":"1016569","partNo":"0102AAG04700N","partDescription":"ASSY ENGINE COMP-RH-0102AAG04700N-W105","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"6","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":64,"assyCategory":"BIW","sapCode":"1016570","partNo":"0102AAG04710N","partDescription":"ASSY ENGINE COMP-LH-0102AAG04710N-W105","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"5.5","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":65,"assyCategory":"BIW","sapCode":"1015712","partNo":"0101BAG05230N","partDescription":"SILL SIDE ASSY INNER RH-0101BAG05230N","model":"SCORPIO","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4.5","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":66,"assyCategory":"BIW","sapCode":"1015713","partNo":"0101BAG05240N","partDescription":"SILL SIDE ASSY INNER LH-0101BAG05240N","model":"SCORPIO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"9","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":67,"assyCategory":"BIW","sapCode":"1004576","partNo":"0101DG0420N","partDescription":"ASSY ROOF BOW FRT","model":"SCORPIO + SC (2) + DC (2)","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"6","cycleTime":"2.2","jph":"22.9","capacityPerDay":"510.0"},{"srNo":68,"assyCategory":"BIW","sapCode":"1004577","partNo":"0101DG0430N","partDescription":"ROOF BOW INTERMEDIATE SPLIT A.C.","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"1","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":69,"assyCategory":"BIW","sapCode":"1004585","partNo":"0103AG0150N","partDescription":"REINF ASSY FRT DR INTRUSION-0103AG0150N","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"2.2","jph":"22.9","capacityPerDay":"510.0"},{"srNo":70,"assyCategory":"BIW","sapCode":"1004556","partNo":"0103BG0220N RH","partDescription":"REINF ASSY RR DR INTRUSION RH","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"7","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":71,"assyCategory":"BIW","sapCode":"1004557","partNo":"0103BG0230N LH","partDescription":"REINF ASSY RR DR INTRUSON LH-0103BG0230N","model":"SCORPIO + SC + DC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"2","cycleTime":"4.5","jph":"11.5","capacityPerDay":"255.0"},{"srNo":72,"assyCategory":"BIW","sapCode":"1030635","partNo":"0102AAL00150N LHD","partDescription":"FG-ASSY ENG COMPRH-LHD0102AAL00150N-W109","model":"SC + DC (LHD)","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"5","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":73,"assyCategory":"BIW","sapCode":"1030636","partNo":"0102AAL00160N LHD","partDescription":"FG-ASSY ENG COMPLH-LHD0102AAL00160N-W109","model":"SC + DC (LHD)","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"6","cycleTime":"6.8","jph":"7.5","capacityPerDay":"166.9"},{"srNo":74,"assyCategory":"BIW","sapCode":"1004572","partNo":"0101BL0020N SC","partDescription":"SILL SIDE INNER SC RH","model":"SC","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":75,"assyCategory":"BIW","sapCode":"1004573","partNo":"0101BL0030N SC","partDescription":"ASSY SILL SIDE INNER SC LH-0101BL0030N","model":"SC","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"5","cycleTime":"14.0","jph":"3.6","capacityPerDay":"81.1"},{"srNo":76,"assyCategory":"BIW","sapCode":"1004574","partNo":"0101BL0190N DC","partDescription":"SILL SIDE INNER DC RH","model":"DC","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"5.6","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":77,"assyCategory":"BIW","sapCode":"1004575","partNo":"0101BL0200N DC","partDescription":"ASSY SILL SIDE INNER DC LH-0101BL0200N","model":"DC","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":78,"assyCategory":"TCF","sapCode":"1066909","partNo":"0401CEE00220N","partDescription":"RTB 220N","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":79,"assyCategory":"TCF","sapCode":"1068660","partNo":"0401CEE00300N","partDescription":"RTB 300N EV","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":80,"assyCategory":"TCF","sapCode":"1092602","partNo":"0401CDE00130N","partDescription":"RTB 130N EPB","model":"XUV 3XO -EPB","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":81,"assyCategory":"TCF","sapCode":"1044223","partNo":"0201ABA00030N","partDescription":"KFRAME 030N","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":82,"assyCategory":"TCF","sapCode":"1068661","partNo":"0201AAW00070N","partDescription":"EV KFRAME 070N","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":83,"assyCategory":"TCF","sapCode":"1050355","partNo":"0201ABD00050N","partDescription":"GASOLINE KFRAME 050N","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Robotic","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":84,"assyCategory":"TCF","sapCode":"1045298","partNo":"0119AS200090N","partDescription":"BUMPER 200090N","model":"XUV 4OO + EXPORT","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":85,"assyCategory":"TCF","sapCode":"1089843","partNo":"0102AS201160N","partDescription":"BUMPER 201160N (Radar)","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"Robotic","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":86,"assyCategory":"TCF","sapCode":"1066909","partNo":"0401CEE00220N","partDescription":"RTB 220N","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"IT GUN","ownership":"JBM","processType":"Manual","cellNo":"Cell 7","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":87,"assyCategory":"TCF","sapCode":"1068660","partNo":"0401CEE00300N","partDescription":"RTB 300N EV","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":88,"assyCategory":"TCF","sapCode":"1092602","partNo":"0401CDE00130N","partDescription":"RTB 130N EPB","model":"XUV 3XO -EPB","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"Manual","cellNo":"Cell 8 (New)","operation":"OP 40","stdManhead":"4","cycleTime":"19.1","jph":"2.7","capacityPerDay":"59.5"},{"srNo":89,"assyCategory":"TCF","sapCode":"NA","partNo":"SPM-1","partDescription":"SPM-1","model":"XUV 3XO","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"SPM","cellNo":"Cell 6 (Old)","operation":"OP 40","stdManhead":"6","cycleTime":"5.8","jph":"8.8","capacityPerDay":"195.5"},{"srNo":90,"assyCategory":"TCF","sapCode":"NA","partNo":"SPM-2","partDescription":"SPM-2","model":"XUV 4OO","variant":"XUV 400","prodShopArea":"SSW LINE","ownership":"JBM","processType":"SPM","cellNo":"Cell 7","operation":"OP 40","stdManhead":"3","cycleTime":"5.8","jph":"8.8","capacityPerDay":"195.5"},{"srNo":91,"assyCategory":"TCF","sapCode":"NA","partNo":"SPM-3","partDescription":"SPM-3","model":"XUV 3XO -EPB","variant":"XUV 400","prodShopArea":"BIW","ownership":"JBM","processType":"SPM","cellNo":"Cell 7","operation":"OP 40","stdManhead":"2","cycleTime":"5.8","jph":"8.8","capacityPerDay":"195.5"}];

const dropdownOptions = {
    assyCategory: ['BIW', 'TCF'],
    model: ['DC', 'EXPORT', 'LHD', 'OPTION', 'SC', 'SC + DC', 'SC + DC (LHD)', 'SCORPIO', 'SCIRPIO + SC (2) + DC (2)', 'SCORPIO + SC + DC', 'W502', 'XUV 3XO', 'XUV 3XO -EPB', 'XUV 3XO + XUV 4OO', 'XUV 4OO', 'XUV 4OO + EXPORT'],
    variant: ['3XO', 'XUV 400', 'EXPORT', 'LHD', 'W502', 'XUV 3XO'],
    prodShopArea: ['BIW', 'NEW SHED', 'MIG WELDING', 'SSW LINE', 'IT GUN', 'XUV 3XO + XUV 4OO'],
    ownership: ['JBM', 'Annapurna', 'BS Ent', 'Trimurti'],
    processType: ['Robotic', 'Manual', 'Spot', 'SSW', 'SPM'],
    cellNo: ['3', '4 (Old)', '4 (New)', '5', '6 (Old)', '6 (New)', '7', '8 (Old)', '8 (New)', '10', '11', 'MANUAL', 'Cell 3', 'Cell 4 (Old)', 'Cell 4 (New)', 'Cell 5', 'Cell 6 (Old)', 'Cell 6 (New)', 'Cell 7', 'Cell 8 (Old)', 'Cell 8 (New)', 'Cell 10', 'Cell 11'],
    operation: ['Manual', 'OP 10', 'OP 10/ OP 20', 'OP 40', 'OP 20', 'OP10', 'MANUAL', 'OP10/OP20']
};

let allData = [];
let rowCounter = 1;
let currentDate = getTodayUTC();

document.addEventListener('DOMContentLoaded', async () => {
    document.getElementById('dateSelector').value = currentDate;
    document.getElementById('dateSelector').addEventListener('change', handleDateChange);
    await loadData();
    renderTable();
    initEventListeners();
});

function initEventListeners() {
    document.getElementById('btnAddRow').addEventListener('click', addNewRow);
    document.getElementById('btnExportExcel').addEventListener('click', exportToExcel);
    document.getElementById('btnUploadCSV').addEventListener('click', () => {
        document.getElementById('csvUpload').click();
    });
    document.getElementById('csvUpload').addEventListener('change', handleCSVUpload);
}

async function handleDateChange(event) {
    currentDate = event.target.value;
    await loadData();
    renderTable();
}

function getDateCollection(dateStr) {
    return collection(db, 'daily_entries', dateStr, 'data');
}

function getDateMetadataDoc(dateStr) {
    return doc(db, 'daily_entries', dateStr);
}

async function loadData() {
    showLoading();
    try {
        const dateCollectionRef = getDateCollection(currentDate);
        const snapshot = await getDocs(dateCollectionRef);
        
        allData = [];
        snapshot.forEach((docSnap) => {
            const data = docSnap.data();
            allData.push({
                id: docSnap.id,
                ...data,
                date: parseDateUTC(currentDate)
            });
        });
        
        // Always ensure we have 91 rows with master data
        if (allData.length < 91) {
            for (let i = allData.length; i < 91; i++) {
                const masterData = MASTER_DATA[i] || {};
                allData.push({
                    id: '',
                    assyCategory: masterData.assyCategory || '',
                    date: parseDateUTC(currentDate),
                    sapCode: masterData.sapCode || '',
                    partNo: masterData.partNo || '',
                    partDescription: masterData.partDescription || '',
                    model: masterData.model || '',
                    variant: masterData.variant || '',
                    prodShopArea: masterData.prodShopArea || '',
                    ownership: masterData.ownership || '',
                    processType: masterData.processType || '',
                    cellNo: masterData.cellNo || '',
                    operation: masterData.operation || '',
                    stdManhead: parseFloat(masterData.stdManhead) || 0,
                    cycleTime: parseFloat(masterData.cycleTime) || 0,
                    jph: parseFloat(masterData.jph) || 0,
                    capacityPerDay: parseFloat(masterData.capacityPerDay) || 0,
                    targetManPerPart: 0,
                    todaysPlan: 0,
                    reqMandays: 0,
                    mpA: 0,
                    mpB: 0,
                    mpC: 0,
                    totalMdays: 0,
                    aShiftProduction: 0,
                    bShiftProduction: 0,
                    cShiftProduction: 0,
                    totalProduction: 0,
                    qualityOK: 0,
                    rejections: 0,
                    rejectionReason: '',
                    achievedManPerPart: 0,
                    lossStartup: 0,
                    lossSetup: 0,
                    lossFixture: 0,
                    lossHR: 0,
                    lossPress: 0,
                    lossStore: 0,
                    lossQC: 0,
                    lossMaintRobotProg: 0,
                    lossMaintFault: 0,
                    lossMaintShank: 0,
                    lossMaintClamp: 0,
                    lossMaintLogic: 0,
                    lossMaintUtility: 0,
                    lossMaintSensor: 0,
                    lossMaintMig: 0,
                    lossMaintTucker: 0,
                    lossMaintSSW: 0,
                    lossMaintSPM: 0,
                    lossPPC: 0,
                    lossMgmt: 0,
                    totalLossTime: 0,
                    noPlanTime: 0,
                    lossShift: '',
                    lossDuration: '',
                    capacityUtilization: 0,
                    actualManPerPart: 0,
                    availabilityPct: 0,
                    performancePct: 0,
                    qualityPct: 0,
                    oeePct: 0,
                    dateString: currentDate
                });
            }
        }
        
        rowCounter = allData.length + 1;
    } catch (error) {
        console.error('Error loading data:', error);
        alert('Error loading data: ' + error.message);
    }
    hideLoading();
}

async function initializeDefaultRows() {
    showLoading();
    try {
        const dateCollectionRef = getDateCollection(currentDate);
        
        // Create 91 rows in Firestore from master data
        for (let i = 0; i < 91; i++) {
            const masterData = MASTER_DATA[i] || {};
            const rowData = {
                assyCategory: masterData.assyCategory || '',
                date: parseDateUTC(currentDate),
                sapCode: masterData.sapCode || '',
                partNo: masterData.partNo || '',
                partDescription: masterData.partDescription || '',
                model: masterData.model || '',
                variant: masterData.variant || '',
                prodShopArea: masterData.prodShopArea || '',
                ownership: masterData.ownership || '',
                processType: masterData.processType || '',
                cellNo: masterData.cellNo || '',
                operation: masterData.operation || '',
                stdManhead: parseFloat(masterData.stdManhead) || 0,
                cycleTime: parseFloat(masterData.cycleTime) || 0,
                jph: parseFloat(masterData.jph) || 0,
                capacityPerDay: parseFloat(masterData.capacityPerDay) || 0,
                targetManPerPart: 0,
                todaysPlan: 0,
                reqMandays: 0,
                mpA: 0,
                mpB: 0,
                mpC: 0,
                totalMdays: 0,
                aShiftProduction: 0,
                bShiftProduction: 0,
                cShiftProduction: 0,
                totalProduction: 0,
                qualityOK: 0,
                rejections: 0,
                rejectionReason: '',
                achievedManPerPart: 0,
                lossStartup: 0,
                lossSetup: 0,
                lossFixture: 0,
                lossHR: 0,
                lossPress: 0,
                lossStore: 0,
                lossQC: 0,
                lossMaintRobotProg: 0,
                lossMaintFault: 0,
                lossMaintShank: 0,
                lossMaintClamp: 0,
                lossMaintLogic: 0,
                lossMaintUtility: 0,
                lossMaintSensor: 0,
                lossMaintMig: 0,
                lossMaintTucker: 0,
                lossMaintSSW: 0,
                lossMaintSPM: 0,
                lossPPC: 0,
                lossMgmt: 0,
                totalLossTime: 0,
                noPlanTime: 0,
                lossShift: '',
                lossDuration: '',
                capacityUtilization: 0,
                actualManPerPart: 0,
                availabilityPct: 0,
                performancePct: 0,
                qualityPct: 0,
                oeePct: 0,
                dateString: currentDate
            };
            
            await addDoc(dateCollectionRef, rowData);
        }
        
        await updateDateMetadata(currentDate, 91);
    } catch (error) {
        console.error('Error initializing rows:', error);
    }
    hideLoading();
}

async function updateDateMetadata(dateStr, entryCount) {
    try {
        const metadataRef = getDateMetadataDoc(dateStr);
        await setDoc(metadataRef, {
            date: dateStr,
            entryCount: entryCount,
            lastUpdated: Timestamp.now()
        }, { merge: true });
    } catch (error) {
        console.error('Error updating metadata:', error);
    }
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    allData.forEach((row, index) => {
        tbody.appendChild(createTableRow(row, index));
    });
}

function getMasterDataForRow(rowIndex) {
    if (rowIndex >= 0 && rowIndex < MASTER_DATA.length) {
        return MASTER_DATA[rowIndex];
    }
    return {};
}

function createTableRow(data = {}, index = null) {
    const tr = document.createElement('tr');
    const rowId = data.id || '';
    const rowNum = index !== null ? index + 1 : rowCounter++;
    
    const masterData = getMasterDataForRow(index !== null ? index : rowCounter - 2);
    
    const assyCategory = data.assyCategory || masterData.assyCategory || '';
    const sapCode = data.sapCode || masterData.sapCode || '';
    const partNo = data.partNo || masterData.partNo || '';
    const partDescription = data.partDescription || masterData.partDescription || '';
    const model = data.model || masterData.model || '';
    const variant = data.variant || masterData.variant || '';
    const prodShopArea = data.prodShopArea || masterData.prodShopArea || '';
    const ownership = data.ownership || masterData.ownership || '';
    const processType = data.processType || masterData.processType || '';
    const cellNo = data.cellNo || masterData.cellNo || '';
    const operation = data.operation || masterData.operation || '';
    const stdManhead = data.stdManhead || masterData.stdManhead || '';
    const cycleTime = data.cycleTime || masterData.cycleTime || '';
    const jph = data.jph || masterData.jph || '';
    const capacityPerDay = data.capacityPerDay || masterData.capacityPerDay || '';
    
    tr.innerHTML = `
        <td class="fixed-col col-srno">${rowNum}</td>
        <td class="fixed-col col-assy">${createSelect('assyCategory', assyCategory, rowId)}</td>
        <td class="fixed-col col-date">${createInput('date', data.date ? formatDateForInput(data.date) : currentDate, 'date', rowId)}</td>
        <td class="fixed-col col-sap">${createInput('sapCode', sapCode, 'text', rowId)}</td>
        <td class="fixed-col col-partno">${createInput('partNo', partNo, 'text', rowId)}</td>
        <td class="fixed-col col-desc">${createInput('partDescription', partDescription, 'text', rowId)}</td>
        <td class="fixed-col col-model">${createSelect('model', model, rowId)}</td>
        <td class="fixed-col col-variant">${createSelect('variant', variant, rowId)}</td>
        <td class="fixed-col col-shop">${createSelect('prodShopArea', prodShopArea, rowId)}</td>
        <td class="fixed-col col-owner">${createSelect('ownership', ownership, rowId)}</td>
        <td class="fixed-col col-process">${createSelect('processType', processType, rowId)}</td>
        <td class="fixed-col col-cell">${createSelect('cellNo', cellNo, rowId)}</td>
        <td class="fixed-col col-op">${createSelect('operation', operation, rowId)}</td>
        <td class="fixed-col col-stdman">${createInput('stdManhead', stdManhead, 'number', rowId)}</td>
        <td class="fixed-col col-ct">${createInput('cycleTime', cycleTime, 'number', rowId, 0.01)}</td>
        <td class="fixed-col col-jph cell-readonly">${createReadonly('jph', jph, rowId)}</td>
        <td class="fixed-col col-cap cell-readonly">${createReadonly('capacityPerDay', capacityPerDay, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('targetManPerPart', data.targetManPerPart, rowId)}</td>
        <td class="col-narrow">${createInput('todaysPlan', data.todaysPlan, 'number', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('reqMandays', data.reqMandays, rowId)}</td>
        <td class="col-narrow">${createInput('mpA', data.mpA, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('mpB', data.mpB, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('mpC', data.mpC, 'number', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('totalMdays', data.totalMdays, rowId)}</td>
        <td class="col-narrow">${createInput('aShiftProduction', data.aShiftProduction, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('bShiftProduction', data.bShiftProduction, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('cShiftProduction', data.cShiftProduction, 'number', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('totalProduction', data.totalProduction, rowId)}</td>
        <td class="col-narrow">${createInput('qualityOK', data.qualityOK, 'number', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('rejections', data.rejections, rowId)}</td>
        <td class="col-medium">${createInput('rejectionReason', data.rejectionReason, 'text', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('achievedManPerPart', data.achievedManPerPart, rowId)}</td>
        <td class="col-narrow">${createInput('lossStartup', data.lossStartup, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossSetup', data.lossSetup, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossFixture', data.lossFixture, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossHR', data.lossHR, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossPress', data.lossPress, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossStore', data.lossStore, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossQC', data.lossQC, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintRobotProg', data.lossMaintRobotProg, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintFault', data.lossMaintFault, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintShank', data.lossMaintShank, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintClamp', data.lossMaintClamp, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintLogic', data.lossMaintLogic, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintUtility', data.lossMaintUtility, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintSensor', data.lossMaintSensor, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintMig', data.lossMaintMig, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintTucker', data.lossMaintTucker, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintSSW', data.lossMaintSSW, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMaintSPM', data.lossMaintSPM, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossPPC', data.lossPPC, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossMgmt', data.lossMgmt, 'number', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('totalLossTime', data.totalLossTime, rowId)}</td>
        <td class="col-narrow">${createInput('noPlanTime', data.noPlanTime, 'number', rowId)}</td>
        <td class="col-narrow">${createInput('lossShift', data.lossShift, 'text', rowId)}</td>
        <td class="col-narrow">${createInput('lossDuration', data.lossDuration, 'text', rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('capacityUtilization', data.capacityUtilization, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('actualManPerPart', data.actualManPerPart, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('availabilityPct', data.availabilityPct, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('performancePct', data.performancePct, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('qualityPct', data.qualityPct, rowId)}</td>
        <td class="col-narrow cell-readonly">${createReadonly('oeePct', data.oeePct, rowId)}</td>
        <td><button class="btn btn-sm btn-danger" onclick="deleteRow('${rowId}')"><i class="fas fa-trash"></i></button></td>
    `;
    
    attachRowEventListeners(tr, rowId);
    return tr;
}

function createInput(field, value = '', type = 'text', rowId = '', step = null) {
    const val = value !== null && value !== undefined ? value : '';
    const stepAttr = step ? `step="${step}"` : '';
    return `<input type="${type}" class="form-control form-control-sm cell-input" 
            data-field="${field}" data-row="${rowId}" value="${val}" ${stepAttr}>`;
}

function createSelect(field, value = '', rowId = '') {
    const options = dropdownOptions[field] || [];
    let html = `<select class="form-select form-select-sm cell-input" data-field="${field}" data-row="${rowId}">
        <option value=""></option>`;
    options.forEach(opt => {
        html += `<option value="${opt}" ${value === opt ? 'selected' : ''}>${opt}</option>`;
    });
    html += `</select>`;
    return html;
}

function createReadonly(field, value = '', rowId = '') {
    const val = value !== null && value !== undefined ? value : '';
    return `<input type="text" class="form-control form-control-sm cell-input" 
            data-field="${field}" data-row="${rowId}" value="${val}" readonly>`;
}

function attachRowEventListeners(tr, rowId) {
    const inputs = tr.querySelectorAll('.cell-input');
    inputs.forEach(input => {
        input.addEventListener('change', (e) => handleCellChange(e, rowId));
    });
}

async function handleCellChange(event, rowId) {
    const field = event.target.getAttribute('data-field');
    const value = event.target.value;
    
    const rowIndex = allData.findIndex(row => row.id === rowId);
    if (rowIndex === -1) return;
    
    allData[rowIndex][field] = value;
    
    recalculateRow(allData[rowIndex]);
    
    // If row has no ID, save it to Firestore first
    if (!rowId) {
        const dateCollectionRef = getDateCollection(currentDate);
        const docRef = await addDoc(dateCollectionRef, allData[rowIndex]);
        allData[rowIndex].id = docRef.id;
    } else {
        await saveRow(rowId, allData[rowIndex]);
    }
    
    renderTable();
}

function recalculateRow(row) {
    const cycleTime = parseFloat(row.cycleTime) || 0;
    const stdManhead = parseFloat(row.stdManhead) || 0;
    
    if (cycleTime > 0) {
        row.jph = ((60 / cycleTime) * EFFICIENCY).toFixed(1);
        row.capacityPerDay = ((WORKING_MINUTES / cycleTime) * EFFICIENCY).toFixed(1);
    }
    
    if (stdManhead > 0 && cycleTime > 0) {
        row.targetManPerPart = (stdManhead * (cycleTime / 60)).toFixed(2);
    }
    
    const todaysPlan = parseFloat(row.todaysPlan) || 0;
    if (todaysPlan > 0 && stdManhead > 0 && cycleTime > 0) {
        row.reqMandays = ((todaysPlan * stdManhead * cycleTime) / (60 * 8)).toFixed(2);
    }
    
    const mpA = parseFloat(row.mpA) || 0;
    const mpB = parseFloat(row.mpB) || 0;
    const mpC = parseFloat(row.mpC) || 0;
    row.totalMdays = (mpA + mpB + mpC).toFixed(0);
    
    const aShiftProduction = parseFloat(row.aShiftProduction) || 0;
    const bShiftProduction = parseFloat(row.bShiftProduction) || 0;
    const cShiftProduction = parseFloat(row.cShiftProduction) || 0;
    row.totalProduction = (aShiftProduction + bShiftProduction + cShiftProduction).toFixed(0);
    
    const qualityOK = parseFloat(row.qualityOK) || 0;
    const totalProduction = parseFloat(row.totalProduction) || 0;
    row.rejections = (totalProduction - qualityOK).toFixed(0);
    
    if (totalProduction > 0 && cycleTime > 0 && parseFloat(row.totalMdays) > 0) {
        row.achievedManPerPart = ((parseFloat(row.totalMdays) * 8 * 60) / (totalProduction * cycleTime)).toFixed(2);
    }
    
    const lossStartup = parseFloat(row.lossStartup) || 0;
    const lossSetup = parseFloat(row.lossSetup) || 0;
    const lossFixture = parseFloat(row.lossFixture) || 0;
    const lossHR = parseFloat(row.lossHR) || 0;
    const lossPress = parseFloat(row.lossPress) || 0;
    const lossStore = parseFloat(row.lossStore) || 0;
    const lossQC = parseFloat(row.lossQC) || 0;
    const lossMaintRobotProg = parseFloat(row.lossMaintRobotProg) || 0;
    const lossMaintFault = parseFloat(row.lossMaintFault) || 0;
    const lossMaintShank = parseFloat(row.lossMaintShank) || 0;
    const lossMaintClamp = parseFloat(row.lossMaintClamp) || 0;
    const lossMaintLogic = parseFloat(row.lossMaintLogic) || 0;
    const lossMaintUtility = parseFloat(row.lossMaintUtility) || 0;
    const lossMaintSensor = parseFloat(row.lossMaintSensor) || 0;
    const lossMaintMig = parseFloat(row.lossMaintMig) || 0;
    const lossMaintTucker = parseFloat(row.lossMaintTucker) || 0;
    const lossMaintSSW = parseFloat(row.lossMaintSSW) || 0;
    const lossMaintSPM = parseFloat(row.lossMaintSPM) || 0;
    const lossPPC = parseFloat(row.lossPPC) || 0;
    const lossMgmt = parseFloat(row.lossMgmt) || 0;
    
    row.totalLossTime = (lossStartup + lossSetup + lossFixture + lossHR + lossPress + 
                        lossStore + lossQC + lossMaintRobotProg + lossMaintFault + 
                        lossMaintShank + lossMaintClamp + lossMaintLogic + lossMaintUtility + 
                        lossMaintSensor + lossMaintMig + lossMaintTucker + lossMaintSSW + 
                        lossMaintSPM + lossPPC + lossMgmt).toFixed(0);
    
    const capacityPerDay = parseFloat(row.capacityPerDay) || 0;
    if (capacityPerDay > 0) {
        row.capacityUtilization = ((totalProduction / capacityPerDay) * 100).toFixed(1);
    }
    
    if (totalProduction > 0 && parseFloat(row.totalMdays) > 0) {
        row.actualManPerPart = ((parseFloat(row.totalMdays) * 8) / totalProduction).toFixed(2);
    }
    
    const totalLossTime = parseFloat(row.totalLossTime) || 0;
    const noPlanTime = parseFloat(row.noPlanTime) || 0;
    const availableTime = WORKING_MINUTES - totalLossTime - noPlanTime;
    if (availableTime > 0) {
        row.availabilityPct = ((availableTime / WORKING_MINUTES) * 100).toFixed(1);
    }
    
    if (totalProduction > 0 && cycleTime > 0 && availableTime > 0) {
        row.performancePct = (((totalProduction * cycleTime) / availableTime) * 100).toFixed(1);
    }
    
    if (totalProduction > 0) {
        row.qualityPct = ((qualityOK / totalProduction) * 100).toFixed(1);
    }
    
    const availability = parseFloat(row.availabilityPct) || 0;
    const performance = parseFloat(row.performancePct) || 0;
    const quality = parseFloat(row.qualityPct) || 0;
    row.oeePct = ((availability * performance * quality) / 10000).toFixed(1);
}

async function addNewRow() {
    showLoading();
    try {
        const newRowIndex = allData.length;
        const masterData = getMasterDataForRow(newRowIndex);
        
        const newRow = {
            assyCategory: masterData.assyCategory || '',
            date: parseDateUTC(currentDate),
            sapCode: masterData.sapCode || '',
            partNo: masterData.partNo || '',
            partDescription: masterData.partDescription || '',
            model: masterData.model || '',
            variant: masterData.variant || '',
            prodShopArea: masterData.prodShopArea || '',
            ownership: masterData.ownership || '',
            processType: masterData.processType || '',
            cellNo: masterData.cellNo || '',
            operation: masterData.operation || '',
            stdManhead: parseFloat(masterData.stdManhead) || 0,
            cycleTime: parseFloat(masterData.cycleTime) || 0,
            jph: parseFloat(masterData.jph) || 0,
            capacityPerDay: parseFloat(masterData.capacityPerDay) || 0,
            targetManPerPart: 0,
            todaysPlan: 0,
            reqMandays: 0,
            mpA: 0,
            mpB: 0,
            mpC: 0,
            totalMdays: 0,
            aShiftProduction: 0,
            bShiftProduction: 0,
            cShiftProduction: 0,
            totalProduction: 0,
            qualityOK: 0,
            rejections: 0,
            rejectionReason: '',
            achievedManPerPart: 0,
            lossStartup: 0,
            lossSetup: 0,
            lossFixture: 0,
            lossHR: 0,
            lossPress: 0,
            lossStore: 0,
            lossQC: 0,
            lossMaintRobotProg: 0,
            lossMaintFault: 0,
            lossMaintShank: 0,
            lossMaintClamp: 0,
            lossMaintLogic: 0,
            lossMaintUtility: 0,
            lossMaintSensor: 0,
            lossMaintMig: 0,
            lossMaintTucker: 0,
            lossMaintSSW: 0,
            lossMaintSPM: 0,
            lossPPC: 0,
            lossMgmt: 0,
            totalLossTime: 0,
            noPlanTime: 0,
            lossShift: '',
            lossDuration: '',
            capacityUtilization: 0,
            actualManPerPart: 0,
            availabilityPct: 0,
            performancePct: 0,
            qualityPct: 0,
            oeePct: 0,
            dateString: currentDate
        };
        
        const dateCollectionRef = getDateCollection(currentDate);
        const docRef = await addDoc(dateCollectionRef, newRow);
        
        newRow.id = docRef.id;
        allData.push(newRow);
        
        await updateDateMetadata(currentDate, allData.length);
        
        renderTable();
    } catch (error) {
        console.error('Error adding row:', error);
        alert('Error adding row: ' + error.message);
    }
    hideLoading();
}

async function saveRow(rowId, rowData) {
    try {
        const dateCollectionRef = getDateCollection(currentDate);
        const docRef = doc(dateCollectionRef, rowId);
        await updateDoc(docRef, rowData);
    } catch (error) {
        console.error('Error saving row:', error);
    }
}

async function deleteRow(rowId) {
    if (!confirm('Are you sure you want to delete this row?')) return;
    
    showLoading();
    try {
        const dateCollectionRef = getDateCollection(currentDate);
        await deleteDoc(doc(dateCollectionRef, rowId));
        
        allData = allData.filter(row => row.id !== rowId);
        await updateDateMetadata(currentDate, allData.length);
        
        renderTable();
    } catch (error) {
        console.error('Error deleting row:', error);
        alert('Error deleting row: ' + error.message);
    }
    hideLoading();
}

window.deleteRow = deleteRow;

function getTodayUTC() {
    const now = new Date();
    return now.toISOString().split('T')[0];
}

function parseDateUTC(dateStr) {
    return new Date(dateStr + 'T00:00:00Z');
}

function formatDateForInput(date) {
    if (!date) return '';
    if (typeof date === 'string') return date.split('T')[0];
    return date.toISOString().split('T')[0];
}

function parseNumber(val) {
    const num = parseFloat(val);
    return isNaN(num) ? 0 : num;
}

async function handleCSVUpload(event) {
    showLoading();
    try {
        const file = event.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        const rows = text.split('\n').map(row => parseCSVRow(row)).filter(row => row.length > 0);
        
        if (rows.length < 2) {
            alert('CSV file is empty or invalid');
            hideLoading();
            return;
        }
        
        const uploadDate = currentDate;
        const dateCollectionRef = getDateCollection(uploadDate);
        
        // Clear existing data for this date
        const existingSnapshot = await getDocs(dateCollectionRef);
        const deletePromises = [];
        existingSnapshot.forEach((doc) => {
            deletePromises.push(deleteDoc(doc.ref));
        });
        await Promise.all(deletePromises);
        
        allData = [];
        let importedCount = 0;
        
        // Process CSV rows
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i];
            if (!cells[0] || cells[0].trim() === '') continue;
            
            const srNo = parseInt(cells[0]);
            const masterData = MASTER_DATA.find(m => m.srNo === srNo) || {};
            
            // Use master data for fixed columns, CSV data for other columns
            const rowData = {
                // Fixed columns from master data
                assyCategory: masterData.assyCategory || '',
                sapCode: masterData.sapCode || '',
                partNo: masterData.partNo || '',
                partDescription: masterData.partDescription || '',
                model: masterData.model || '',
                variant: masterData.variant || '',
                prodShopArea: masterData.prodShopArea || '',
                ownership: masterData.ownership || '',
                processType: masterData.processType || '',
                cellNo: masterData.cellNo || '',
                operation: masterData.operation || '',
                stdManhead: parseFloat(masterData.stdManhead) || 0,
                cycleTime: parseFloat(masterData.cycleTime) || 0,
                jph: parseFloat(masterData.jph) || 0,
                capacityPerDay: parseFloat(masterData.capacityPerDay) || 0,
                
                // Date from current selection
                date: parseDateUTC(uploadDate),
                
                // Editable columns from CSV (columns 18 onwards)
                targetManPerPart: parseNumber(cells[17]),
                todaysPlan: parseNumber(cells[18]),
                reqMandays: parseNumber(cells[19]),
                mpA: parseNumber(cells[20]),
                mpB: parseNumber(cells[21]),
                mpC: parseNumber(cells[22]),
                totalMdays: parseNumber(cells[23]),
                aShiftProduction: parseNumber(cells[24]),
                bShiftProduction: parseNumber(cells[25]),
                cShiftProduction: parseNumber(cells[26]),
                totalProduction: parseNumber(cells[27]),
                qualityOK: parseNumber(cells[28]),
                rejections: parseNumber(cells[29]),
                rejectionReason: cells[30] || '',
                achievedManPerPart: parseNumber(cells[31]),
                lossStartup: parseNumber(cells[32]),
                lossSetup: parseNumber(cells[33]),
                lossFixture: parseNumber(cells[34]),
                lossHR: parseNumber(cells[35]),
                lossPress: parseNumber(cells[36]),
                lossStore: parseNumber(cells[37]),
                lossQC: parseNumber(cells[38]),
                lossMaintRobotProg: parseNumber(cells[39]),
                lossMaintFault: parseNumber(cells[40]),
                lossMaintShank: parseNumber(cells[41]),
                lossMaintClamp: parseNumber(cells[42]),
                lossMaintLogic: parseNumber(cells[43]),
                lossMaintUtility: parseNumber(cells[44]),
                lossMaintSensor: parseNumber(cells[45]),
                lossMaintMig: parseNumber(cells[46]),
                lossMaintTucker: parseNumber(cells[47]),
                lossMaintSSW: parseNumber(cells[48]),
                lossMaintSPM: parseNumber(cells[49]),
                lossPPC: parseNumber(cells[50]),
                lossMgmt: parseNumber(cells[51]),
                totalLossTime: parseNumber(cells[52]),
                noPlanTime: parseNumber(cells[53]),
                lossShift: cells[54] || '',
                lossDuration: cells[55] || '',
                capacityUtilization: parseNumber(cells[56]),
                actualManPerPart: parseNumber(cells[57]),
                availabilityPct: parseNumber(cells[58]),
                performancePct: parseNumber(cells[59]),
                qualityPct: parseNumber(cells[60]),
                oeePct: parseNumber(cells[61]),
                dateString: uploadDate
            };
            
            const docRef = await addDoc(dateCollectionRef, rowData);
            importedCount++;
            
            allData.push({
                id: docRef.id,
                ...rowData,
                date: parseDateUTC(uploadDate)
            });
        }
        
        // Fill remaining rows up to 91 with master data only
        for (let i = importedCount; i < 91; i++) {
            const masterData = MASTER_DATA[i] || {};
            const rowData = {
                assyCategory: masterData.assyCategory || '',
                date: parseDateUTC(uploadDate),
                sapCode: masterData.sapCode || '',
                partNo: masterData.partNo || '',
                partDescription: masterData.partDescription || '',
                model: masterData.model || '',
                variant: masterData.variant || '',
                prodShopArea: masterData.prodShopArea || '',
                ownership: masterData.ownership || '',
                processType: masterData.processType || '',
                cellNo: masterData.cellNo || '',
                operation: masterData.operation || '',
                stdManhead: parseFloat(masterData.stdManhead) || 0,
                cycleTime: parseFloat(masterData.cycleTime) || 0,
                jph: parseFloat(masterData.jph) || 0,
                capacityPerDay: parseFloat(masterData.capacityPerDay) || 0,
                targetManPerPart: 0,
                todaysPlan: 0,
                reqMandays: 0,
                mpA: 0,
                mpB: 0,
                mpC: 0,
                totalMdays: 0,
                aShiftProduction: 0,
                bShiftProduction: 0,
                cShiftProduction: 0,
                totalProduction: 0,
                qualityOK: 0,
                rejections: 0,
                rejectionReason: '',
                achievedManPerPart: 0,
                lossStartup: 0,
                lossSetup: 0,
                lossFixture: 0,
                lossHR: 0,
                lossPress: 0,
                lossStore: 0,
                lossQC: 0,
                lossMaintRobotProg: 0,
                lossMaintFault: 0,
                lossMaintShank: 0,
                lossMaintClamp: 0,
                lossMaintLogic: 0,
                lossMaintUtility: 0,
                lossMaintSensor: 0,
                lossMaintMig: 0,
                lossMaintTucker: 0,
                lossMaintSSW: 0,
                lossMaintSPM: 0,
                lossPPC: 0,
                lossMgmt: 0,
                totalLossTime: 0,
                noPlanTime: 0,
                lossShift: '',
                lossDuration: '',
                capacityUtilization: 0,
                actualManPerPart: 0,
                availabilityPct: 0,
                performancePct: 0,
                qualityPct: 0,
                oeePct: 0,
                dateString: uploadDate
            };
            
            const docRef = await addDoc(dateCollectionRef, rowData);
            
            allData.push({
                id: docRef.id,
                ...rowData,
                date: parseDateUTC(uploadDate)
            });
        }
        
        await updateDateMetadata(uploadDate, allData.length);
        renderTable();
        
        alert(`Successfully imported ${importedCount} rows from CSV and filled remaining ${91 - importedCount} rows with master data`);
        
    } catch (error) {
        console.error('Error processing CSV:', error);
        alert('Error processing CSV: ' + error.message);
    } finally {
        event.target.value = '';
        hideLoading();
    }
}

function parseCSVRow(row) {
    const cells = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < row.length; i++) {
        const char = row[i];
        const nextChar = row[i + 1];
        
        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            cells.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    cells.push(current.trim());
    return cells;
}

function exportToExcel() {
    const data = allData.map((item, index) => ({
        'Sr. No.': index + 1,
        'Assy Cat.': item.assyCategory,
        'Date': item.dateString || formatDateForInput(item.date),
        'SAP Code': item.sapCode,
        'Part No.': item.partNo,
        'Part Description': item.partDescription,
        'Model': item.model,
        'Variant': item.variant,
        'Prod Shop/Area': item.prodShopArea,
        'Ownership': item.ownership,
        'Process Type': item.processType,
        'Cell No.': item.cellNo,
        'OP': item.operation,
        'Std. ManHead': item.stdManhead,
        'CT (Min.)': item.cycleTime,
        'JPH': item.jph,
        'Cap/day': item.capacityPerDay,
        'Tgt Man/part': item.targetManPerPart,
        'Plan': item.todaysPlan,
        'Req. mdays': item.reqMandays,
        'A Shift MP': item.mpA,
        'B Shift MP': item.mpB,
        'C Shift MP': item.mpC,
        'Total Mdays': item.totalMdays,
        'A Prod': item.aShiftProduction,
        'B Prod': item.bShiftProduction,
        'C Prod': item.cShiftProduction,
        'Total Prod': item.totalProduction,
        'Q. OK': item.qualityOK,
        'NG/Rej': item.rejections,
        'Reason': item.rejectionReason,
        'Ach Man/Part': item.achievedManPerPart,
        'Startup': item.lossStartup,
        'Setup Chg': item.lossSetup,
        'Fixture': item.lossFixture,
        'HR Loss': item.lossHR,
        'Press': item.lossPress,
        'BOP': item.lossStore,
        'QC': item.lossQC,
        'Robot Prog': item.lossMaintRobotProg,
        'Robot Fault': item.lossMaintFault,
        'Shank': item.lossMaintShank,
        'Clamp': item.lossMaintClamp,
        'Thickness': item.lossMaintLogic,
        'Air/Gas': item.lossMaintUtility,
        'Sensor': item.lossMaintSensor,
        'Mig': item.lossMaintMig,
        'Tucker': item.lossMaintTucker,
        'SSW': item.lossMaintSSW,
        'SPM': item.lossMaintSPM,
        'PPC': item.lossPPC,
        'Mgmt': item.lossMgmt,
        'Total Loss': item.totalLossTime,
        'No Plans': item.noPlanTime,
        'Shift': item.lossShift,
        'Duration': item.lossDuration,
        'Cap Util%': item.capacityUtilization,
        'Act Man/Part': item.actualManPerPart,
        'Avail%': item.availabilityPct,
        'Perf%': item.performancePct,
        'Qual%': item.qualityPct,
        'OEE%': item.oeePct
    }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MIS Data');
    XLSX.writeFile(wb, `JBM_MIS_Data_${currentDate}.xlsx`);
}

function showLoading() {
    document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
}