from pubchempy import Compound, get_compounds
from multiprocessing import Pool
import xlrd
import math
import random
a=['(Z)-Luteoxanthin','apigenin','apigenin 7-glucuronide','asiatic acid','Asperuloside acid','baicalein','baicalin','beta caryophyllene','beta sitosterol','beta-amyrin','beta-pinene','betulinic acid','betulinol','biochanin','Biochanin A','borreline','Brazilin','b-sitosterol','butein','calycosin','campesterol','caryophyllene','Catechin','chlorogenic acid','Choline','Citric acid','copaene','Coumarin','coumarins','crocetin dimethyl ester','cryptoxanthin','daidzein','dicarboxylic acid','Dihydrogallic acid','diosgenin','d-pinitol','Ellagic acid','epicatechin','epigallocatechin','epigallocatechin-3-gallate','erythrodiol','esculetin','flavesone','formononetin','gallic acid','gallocatechin','genistein','gentisic acid','germacrene D','hematoxylol A','humulene','Hydroxycinnamic acid','hydroxyl benzoic acid','ichangin','indene','isolariciresinol','Isorhamnetin','Kaempferol','laballenic acid','lariciresinol','leptospermone','linalool','linoleic acid','liquiritigenin','L-tryptophan','Lupeol','luteolin','luteolin 7-O-glucuronide','mangiferin','medioresinol','methyl gallate','methylglyoxal','momordenol','momordicilin','momordicinin','myrcene','myricetin','naringenin','naringin','nobiletin','n-propyl gallate','obacunone','oleanolic acid','oleic acid','orientin','Palmitic acid','palmitic acid methyl ester','phloridzin','Phyllanthol','phytol','pipecolic acid','pratensein','procyanidins','propterol','prunetin','Pterostilben','quercetin','resveratrol','robinetin','rutin','scandoside','sesamin','silymarin ','squalene','stigmasterol','syringin','tarphetalin','terpene alcohol ','terpineol','ursolic acid','verticillatine A','vitamin c]','m-digallic acid','protocatechuic acid','methyl ethyl ester','mucilage','saponin','Galactose','arabinose','glucuronic acid','4-O-methylglucuronic acid  rhamnose','stearic acid','citronellol','phthalic acid ','rhamnetin','cyanidin','cardanol delphinidin','2-hydroxy- 6-pentadecylbenzoic acid','salicylic acid  ','caffeic acid','p-hydroxy benzoic acid','ferulic acid','vanilic acid','syringic acid','p-coumaric acid','sinapic acid','n-hentriacontane','avicennone A','2-caffeoyl-mussaenosidic acid','2-CoU mamaheswarraoroyl-mussaenosidic acid','Azadirachtol','azadirachnol','nimocinol','nimocinolide','nimbocinone','nimolinone','isonimocinolide','nimocin','nimbocetin','nimbochalcin ','Asperolosidic acid','scoside','borrecoxine','Casearins A-X','gallic acid derivatives','Kaempferein','isorhein','chrysophenol','imodin','aloe-imodin','thein','sannasoides','stetculic acids','procyanidin B2','biflavonoids','chrysophanol','vernolic','malvalic','sennoside B','triflavonoids','rhein','rhein glucoside','Tannins','resins','fats  oils','glycosides ','β-sitosterol','5- deoxypulchelloside','cirsimaritin 4-O-β-D-glucopyranoside ','4- sodium sulphate cirsimaritin 4-O-β-D-glucopyranoside ','Kampferol','limonin','limonoids','alpha-tocopherol','lauric acid ','vitamin C','l-arginine','isoskimmiwallin','skimmiwallin ','Cardiaquinone A','Cardiaquinone B','Cardiaquinone C','Cardiaquinone D','4-methyl 4-ethenyl-3-(1-methyl ethenyl)-1-(1-methyl methanol)cyclohexane','β-eudesmol','spathulenol','meroterpenoid ','quinones','cadina','alpha-amyrin','octacosanol','lupeol-3-rhamnoside','taxifolin-3','5-dirhamnoside','betulin ','beta-sitosterol','hentricontanol','hentricontane','hesperitin-7-rhamnoside ','Emodin','physcion','royleanone','α-amyrin  β-sitosterol','sinapic acids','isoquercitrine ','Coumaric acid','stigmastanol  glucoside','melilotic acid','isorhamnetin ','succinic acid','fumaric acid','p-hydroxybenzoic acid ','4-hydroxy isophthalic acid','3 4-dihydroxycinnamic acid ','esculetin ','isowedelolactone','Gouanoside A','gouanoside','Guaianin  guaiacin','protossapanin A','protossapanin B','hematoxin','methylhematoxylol','hematoxylol B','hematoxylin','10-O-methylepihematoxylol B','3-deoxysappanchalcone','β-caryophyllene','thymol','p-cymene','sabinene','sesquiterpenes','caryophyllene oxide','calamusenone','geraniol','phytone','nonacosane','(Z)-hex-3-enyl benzoate','α-terpineol','γ-eudesmol','helioxanthin','(+)-isolariciresinol  justicinol','aryltetralin','(+)-lariciresinol','alphitolic acid','4- hydroxyl benzoic acid','23-hydroxyursolic acid','nepetaefolinol','labdanic acid','lenotinin','5-methoxy7-hydroxy-6 8-dimethylflavone','morolic acid','2β-acetoxy-3-acetyl morolic acid','elemene','α-pinene','5 7-dimethoxy -6-methylflavone','5-hydroxy-6-methyl-7-methoxyflavone','δ-cadinene','linolenic acid ester','protocatechic acid','4-phenyl gallate','6-phenyl-n-hexyl gallate','phloretin','Jasomonic acid','m-mimosine','d-glucoronic acid','d-xylose','momorcharins','momordicins','momordin','momordolol','charantin','charine','cucurbitins','cucurbitacins','cucurbitanes','cycloartenols','elaeostearic acids','galacturonic acids','goyaglycosides','goyasaponins','multiflorenol','all-E-luteoxanthin','13-Z-lutein','all-E-zeaxanthin','15-Z-β-carotene','glochidone','β-amyrin','lupanyl acetate','friedelin','piperidine alkaloids','juliflorine  (-)-laricirenol','isoliquritigenin','carpucin','propterol-B','alkaloid  resin 5','4-dimethoxy-8-methylisoflavone','4 6 4-trihydroxyaurone 6-O-rhamnopyranoside  ','4 6 4- trihydroxy-7-methylaurone 4-O-rhamnopyranoside','Cadinene','stigmasta-9-en-3 6 7-triol','3-hydroxy-22-epoxystigmastane','phenolic acids','6-Hydroxyluteolol 7-glucuronide','apigenol 7-glucuronide','vitamin  B3','saponins','C-flavonoid glycoside','ellagitannin acid','ellagic acid dehydrate','apigenin 6-c-(2 -galloyl)-L-D glycoside','punicalin','punicalagin','tannin','aempferol','(−)-epicatechin','(+)-catechin','(−)-epigallocatechin','(+)- gallocatechin','(−)-epigallocatechin-3-gallate','silymarin  ','(−)-epicatechin-3-gallate','12-oleanen-3-ol-3ß-acetate','ß-sitosterol','3-𝛽-D-glucopyranosyl-1-hydroxy-6(E)-tetradecene-8 10 12- triyne','2-𝛽-D-glucopyranosyloxy-1-hydroxy-5(E)-tridecene-7 9 11-triyne ','2-𝛽-D-glucopyranosyloxy-1- hydroxytrideca-5 7 9 11-tetrayne cytopiloyne','4 5-Di-O-caffeoylquinic acid','3 5-Di-Ocaffeoylquinic acid','3 4-Di-O-caffeoylquinic acid','alkaloids','steroids','phenolic compounds','cardiac glycosides','sesquiterpene','flavonoid','Urolignosid']
def abc(b):
    found_list = {}
    try:
        for compound in get_compounds(b, 'name'):
            mol=Compound.from_cid(compound.cid).molecular_weight
            wb = xlrd.open_workbook('GCMASS_.xlsx')
            p = wb.sheet_names()
            for y in p:
                sheet_data = []   
                sh = wb.sheet_by_name(y)
                for rownum in range(sh.nrows):
                    sheet_data.append((sh.row_values(rownum)))
                #print(sheet_data[1])
                
                cho=[]
                for i in range(len(sheet_data)):
                    for j in range(len(sheet_data[i])):
                        #print(j)
                        try:
                            if sheet_data[i][j]>math.floor(mol) and sheet_data[i][j]<math.floor(mol)+1 and sheet_data[i][j+1]!=0:
                                
                                cho.append(sheet_data[i][j+1])
                        except:
                            continue
                if len(cho)!=0:
                    found_list[y]=random.choice(cho)
                else:
                    found_list[y]=None
                
            #print(found_list.values(),found_list.keys())
            #print(str(b)+','+str(found_list.get('bp2'))+','+str(found_list.get('bp8'))+','+str(found_list.get('bp10'))+','+str(mol))
            print(str(b)+','+str(found_list.get('hs1'))+','+str(found_list.get('hs2'))+','+str(found_list.get('hs3'))+','+str(found_list.get('hs4'))+','+str(found_list.get('hs5'))+','+str(found_list.get('hs6'))+','+str(found_list.get('hs7'))+','+str(found_list.get('hs8'))+','+str(found_list.get('hs9'))+','+str(found_list.get('hs10'))+','+str(found_list.get('hs11'))+','+str(found_list.get('hs12'))+','+str(found_list.get('hs13'))+','+str(found_list.get('hs14'))+','+str(found_list.get('hs15'))+','+str(found_list.get('hs16'))+','+str(found_list.get('hs17'))+','+str(found_list.get('hs18'))+','+str(found_list.get('hs19'))+','+str(found_list.get('hs20'))+','+str(found_list.get('hs21'))+','+str(found_list.get('hs22'))+','+str(mol))

            return None
    except:
        return None
if __name__ == '__main__':
    print('compound,hs1,hs2,hs3,hs4,hs5,hs6,hs7,hs8,hs9,hs10,hs11,hs12,hs13,hs14,hs15,hs16,hs17,hs18,hs19,hs20,hs21,hs22,MW')
    with Pool(12) as p:
        p.map(abc,a)
    p.close()   