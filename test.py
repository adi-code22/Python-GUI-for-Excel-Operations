window.mainloop()
window = Tk()
window.geometry("500x500")
window.title('File Explorer')
window.config(background="white")

df = pd.read_excel('test.xlsx', index_col=0)

df['Difference'] = df['Column1'] + df['Column2']

print(df.head())
writer = pd.ExcelWriter('test.xlsx')
df.to_excel(writer)
writer.save()
