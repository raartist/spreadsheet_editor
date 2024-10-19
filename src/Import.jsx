const ImportBtn = ({handleFileChange}) => {
    return (
        <input type="file" accept=".xlsx" onChange={handleFileChange} className="import-btn"/>
    )
}
export default ImportBtn;