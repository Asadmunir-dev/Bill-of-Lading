import {
  collection,
  doc,
  setDoc,
  getDocs,
  deleteDoc,
  onSnapshot,
  query,
} from 'firebase/firestore';
import { db, handleFirestoreError, OperationType } from '../firebase';
import { BillOfLadingRecord, UserProfile } from '../types';

/**
 * Clean ID for Firestore document path (alphanumeric, dash, underscore, dot)
 */
export function sanitizeDocId(rawId: string): string {
  const sanitized = rawId.replace(/[^a-zA-Z0-9_\-\.]/g, '_').trim();
  return sanitized.length > 0 ? sanitized.substring(0, 100) : `bl_${Date.now()}`;
}

/**
 * Saves a Bill of Lading record under users/{userId}/bills_of_lading/{billId}
 */
export async function saveBillToFirestore(
  userId: string,
  record: BillOfLadingRecord
): Promise<void> {
  const docId = sanitizeDocId(record.id || record.blNumber);
  const path = `users/${userId}/bills_of_lading/${docId}`;
  
  const payload: BillOfLadingRecord = {
    ...record,
    id: docId,
    userId,
    updatedAt: new Date().toISOString(),
    savedAt: record.savedAt || new Date().toLocaleString(),
  };

  try {
    const docRef = doc(db, 'users', userId, 'bills_of_lading', docId);
    await setDoc(docRef, payload, { merge: true });
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Subscribes in real-time to a user's saved Bills of Lading
 */
export function subscribeUserBills(
  userId: string,
  onUpdate: (bills: BillOfLadingRecord[]) => void,
  onError?: (err: Error) => void
): () => void {
  const path = `users/${userId}/bills_of_lading`;
  const colRef = collection(db, 'users', userId, 'bills_of_lading');
  const q = query(colRef);

  const unsubscribe = onSnapshot(
    q,
    (snapshot) => {
      const records: BillOfLadingRecord[] = [];
      snapshot.forEach((docSnap) => {
        records.push(docSnap.data() as BillOfLadingRecord);
      });
      // Sort by updatedAt descending
      records.sort((a, b) => {
        const timeA = a.updatedAt ? new Date(a.updatedAt).getTime() : 0;
        const timeB = b.updatedAt ? new Date(b.updatedAt).getTime() : 0;
        return timeB - timeA;
      });
      onUpdate(records);
    },
    (error) => {
      console.error('Subscription error for path', path, error);
      if (onError) onError(error);
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );

  return unsubscribe;
}

/**
 * One-time fetch of all saved Bills of Lading
 */
export async function fetchUserBills(userId: string): Promise<BillOfLadingRecord[]> {
  const path = `users/${userId}/bills_of_lading`;
  try {
    const colRef = collection(db, 'users', userId, 'bills_of_lading');
    const snapshot = await getDocs(colRef);
    const records: BillOfLadingRecord[] = [];
    snapshot.forEach((docSnap) => {
      records.push(docSnap.data() as BillOfLadingRecord);
    });
    records.sort((a, b) => {
      const timeA = a.updatedAt ? new Date(a.updatedAt).getTime() : 0;
      const timeB = b.updatedAt ? new Date(b.updatedAt).getTime() : 0;
      return timeB - timeA;
    });
    return records;
  } catch (error) {
    handleFirestoreError(error, OperationType.GET, path);
  }
}

/**
 * Deletes a saved Bill of Lading from Firestore
 */
export async function deleteBillFromFirestore(
  userId: string,
  billId: string
): Promise<void> {
  const docId = sanitizeDocId(billId);
  const path = `users/${userId}/bills_of_lading/${docId}`;
  try {
    const docRef = doc(db, 'users', userId, 'bills_of_lading', docId);
    await deleteDoc(docRef);
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}

/**
 * Saves or updates user profile metadata in users/{userId}
 */
export async function saveUserProfile(
  userId: string,
  profile: Partial<UserProfile>
): Promise<void> {
  const path = `users/${userId}`;
  try {
    const docRef = doc(db, 'users', userId);
    await setDoc(
      docRef,
      {
        ...profile,
        uid: userId,
        updatedAt: new Date().toISOString(),
      },
      { merge: true }
    );
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}
