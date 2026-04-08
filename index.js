// Reference Board Builder_1
// Entry point for the Miro app icon.

async function init() {
  await miro.board.ui.on("icon:click", async () => {
    await miro.board.ui.openPanel({
      url: "panel.html",
    });
  });
}

init();
