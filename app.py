import streamlit as st
import pandas as pd
import io
import html


st.set_page_config(
    page_title="Generatore HTML WooCommerce",
    layout="wide"
)
# DOWNLOAD FILE ESEMPIO

try:

    with open(
        "esempio_file_input.xlsx",
        "rb"
    ) as file:

        st.download_button(

            label="📥 Scarica Excel esempio",

            data=file,

            file_name="esempio_file_input.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

# ==========================
# CSS DEL TEMPLATE PRODOTTO
# ==========================

CSS = """

<style>

.product-container {
max-width:1100px;
margin:auto;
font-family:Arial, Helvetica, sans-serif;
color:#333;
font-size:15px;
line-height:1.6;
}

.product-container * {
box-sizing:border-box;
}


.title-box {
background:#e28413;
padding:20px;
text-align:center;
}


.title-box h1 {
margin:0;
font-size:30px;
color:white;
}


.title-box p {
color:white;
}


.section {
display:flex;
flex-wrap:wrap;
align-items:center;
margin-top:20px;
}


.image-box {
width:42%;
padding:15px;
text-align:center;
}


.image-box img {
width:100%;
max-width:450px;
height:auto;
}


.text-box {
width:58%;
padding:20px;
}


.product-container h2 {
color:#e28413;
}


.info-box {

background:#fafafa;
border:1px solid #ddd;
padding:20px;
margin-top:20px;

}


.gallery {

display:flex;
flex-wrap:wrap;

}


.gallery img {

width:50%;
padding:6px;

}


.product-container table {

width:100%;
border-collapse:collapse;

}


.product-container td {

border:1px solid #ccc;
padding:8px;

}


.product-container td:first-child {

background:#f5f5f5;
font-weight:bold;

}



@media(max-width:600px){

.section{

display:block;

}


.image-box,
.text-box{

width:100%;

}


.gallery img{

width:100%;

}


.title-box h1{

font-size:22px;

}

}

</style>

"""



# ==========================
# FUNZIONI
# ==========================


def lista_caratteristiche(testo):

    if pd.isna(testo):
        return ""

    html_lista=""

    elementi=str(testo).split(";")

    for elemento in elementi:

        if elemento.strip():

            html_lista += (
                f"<li>{elemento.strip()}</li>"
            )

    return html_lista



def crea_tabella(testo):

    if pd.isna(testo):
        return ""

    tabella=""

    elementi=str(testo).split(";")

    for elemento in elementi:

        if ":" in elemento:

            campo,valore = elemento.split(":",1)

            tabella += f"""

            <tr>
            <td>{campo.strip()}</td>
            <td>{valore.strip()}</td>
            </tr>

            """

    return tabella



def crea_galleria(row):

    risultato=""

    for foto in [
        "FOTO3",
        "FOTO4",
        "FOTO5",
        "FOTO6"
    ]:

        if foto in row:

            valore=row[foto]

            if pd.notna(valore) and valore!="":

                risultato += f"""

                <img src="{valore}">

                """

    return risultato




def genera_html(row):


    html_finale=f"""

{CSS}


<div class="product-container">



<div class="title-box">

<h1>
{row['TITOLO']}
</h1>


<p>
{row['SOTTOTITOLO']}
</p>

</div>




<div class="section">


<div class="image-box">

<img src="{row['FOTO1']}">

</div>



<div class="text-box">

<h2>Descrizione</h2>

<p>
{row['DESCRIZIONE']}
</p>

</div>


</div>






<div class="section info-box">


<div class="text-box">


<h2>
Caratteristiche principali
</h2>


<ul>

{lista_caratteristiche(row['CARATTERISTICHE'])}

</ul>


</div>




<div class="image-box">

<img src="{row['FOTO2']}">

</div>



</div>








<div class="section">


<div class="image-box">


<img src="{row['FOTO3']}">


</div>



<div class="text-box">


<h2>
Dati tecnici
</h2>



<table>


{crea_tabella(row['DATI_TECNICI'])}


</table>


</div>



</div>







<h2 style="text-align:center">

Galleria immagini

</h2>




<div class="gallery">


{crea_galleria(row)}


</div>







<div class="info-box">


<h2>
Contenuto confezione
</h2>



<ul>

<li>

{row['CONTENUTO']}

</li>


</ul>


</div>






<div style="

margin-top:20px;

background:#fff8e1;

border-left:5px solid #e28413;

padding:15px;

">


<strong>Nota:</strong>


<br>

{row['NOTA']}


</div>



</div>

"""


    return html_finale





# ==========================
# STREAMLIT
# ==========================


st.title(
"🛒 Generatore descrizioni WooCommerce da Excel"
)


file_excel = st.file_uploader(
"Carica file Excel prodotti",
type=["xlsx"]
)



if file_excel:


    df=pd.read_excel(file_excel)



    st.success(
        f"{len(df)} prodotti caricati"
    )


    st.dataframe(
        df.head()
    )



    if st.button(
        "⚙️ Genera descrizioni HTML"
    ):


        df["DESCRIZIONE_HTML"] = df.apply(
            genera_html,
            axis=1
        )


        output = io.BytesIO()


        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                index=False,
                sheet_name="Prodotti"
            )



        output.seek(0)



        st.success(
            "File generato correttamente"
        )



        st.download_button(

            label="⬇️ Scarica Excel finale",

            data=output,

            file_name="prodotti_woocommerce_html.xlsx",

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )
