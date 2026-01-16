/* global Word console */

export async function insertText(text: string, searchText: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      // Search for the text in the document (case-insensitive)
      const searchResults = context.document.body.search(searchText, {
        matchCase: false,
        matchWholeWord: false,
      });
      
      context.load(searchResults, "text");
      await context.sync();
      
      if (searchResults.items.length === 0) {
        throw new Error(`Text "${searchText}" not found in document`);
      }
      
      // Get the first occurrence
      const firstMatch = searchResults.items[0];
      
      // Insert the text before the found text
      firstMatch.insertText(text, Word.InsertLocation.before);
      
      await context.sync();
      console.log("Text inserted successfully");
    });
  } catch (error) {
    console.error("Error inserting text:", error);
    throw error;
  }
}
