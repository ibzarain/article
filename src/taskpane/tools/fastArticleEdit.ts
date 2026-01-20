import { DocumentChange } from '../types/changes';

let changeTracker: ((change: DocumentChange) => Promise<void>) | null = null;

export function setFastArticleChangeTracker(tracker: (change: DocumentChange) => Promise<void>) {
  changeTracker = tracker;
}

export function getFastArticleChangeTracker() {
  return changeTracker;
}
