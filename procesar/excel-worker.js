// Web Worker: genera el archivo Excel en un hilo aparte del navegador,
// para que la página del procesador no se congele mientras arma archivos
// grandes (como el Excel de "Toda la Microrred" de un mes completo).
importScripts('xlsx.full.min.js');

self.onmessage = function(e){
  try{
    const { headers, datos, nombreArchivo } = e.data;
    self.postMessage({ tipo: "progreso", mensaje: "Armando la hoja..." });
    const aoa = [headers, ...datos];
    const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Consolidado");
    self.postMessage({ tipo: "progreso", mensaje: "Comprimiendo el archivo..." });
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    self.postMessage({ tipo: "listo", buffer: wbout, nombreArchivo }, [wbout.buffer]);
  } catch(err){
    self.postMessage({ tipo: "error", mensaje: err && err.message ? err.message : String(err) });
  }
};
