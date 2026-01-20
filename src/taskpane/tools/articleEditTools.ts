import { DocumentChange } from '../types/changes';

let changeTracker: ((change: DocumentChange) => Promise<void>) | null = null;

export function setArticleChangeTracker(tracker: (change: DocumentChange) => Promise<void>) {
  changeTracker = tracker;
}

export function getArticleChangeTracker() {
  return changeTracker;
}
