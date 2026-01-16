/* global Word console */

export async function insertText(text: string, searchText: string) {
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
        console.log("Text not found in document");
        return;
      }
      
      // Get the first occurrence
      const firstMatch = searchResults.items[0];
      
      // Insert the text before the found text
      firstMatch.insertText(text, Word.InsertLocation.before);
      
      await context.sync();
      console.log("Text inserted successfully");
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
