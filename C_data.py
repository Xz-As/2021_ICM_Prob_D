import os
import pandas as pd
import numpy as np
import xlrd
import xlsxwriter
import networkx as nx
from matplotlib import pyplot as plt


data_path = r'F:\数模\比赛\2021_ICM_Problem_D_Data\2021_ICM_Problem_D_Data'
file_names = ['data_by_artist.csv', 'data_by_year.csv', 'full_music_data.csv', 'influence_data.csv']
artists = {}


def load_data(file_path = data_path, file_names = file_names):
    follower_net = {}
    artists = {}
    musics = []
    for i in range(4):
        file1 = pd.read_csv(file_path + '\\' + file_names[i])
        #print(file1.keys())
        if i == 1:
            continue

        for j in range(len(file1[file1.keys()[0]])):
            if i == 0:
                artists[str(file1['artist_id'][j])] = {}
                artists[str(file1['artist_id'][j])]['name'] = file1['artist_name'][j]
                artists[str(file1['artist_id'][j])]['fl_id'] = []
                artists[str(file1['artist_id'][j])]['danceability'] = file1['danceability'][j]
                artists[str(file1['artist_id'][j])]['energy'] = file1['energy'][j]
                artists[str(file1['artist_id'][j])]['valence'] = file1['valence'][j]
                artists[str(file1['artist_id'][j])]['tempo'] = file1['tempo'][j]
                artists[str(file1['artist_id'][j])]['loudness'] = file1['loudness'][j]
                artists[str(file1['artist_id'][j])]['mode'] = file1['mode'][j]
                artists[str(file1['artist_id'][j])]['key'] = file1['key'][j]
                artists[str(file1['artist_id'][j])]['acousticness'] = file1['acousticness'][j]
                artists[str(file1['artist_id'][j])]['instrumentalness'] = file1['instrumentalness'][j]
                artists[str(file1['artist_id'][j])]['liveness'] = file1['liveness'][j]
                artists[str(file1['artist_id'][j])]['speechiness'] = file1['speechiness'][j]
                artists[str(file1['artist_id'][j])]['duration_ms'] = file1['duration_ms'][j]
                artists[str(file1['artist_id'][j])]['popularity'] = file1['popularity'][j]
                artists[str(file1['artist_id'][j])]['count'] = file1['count'][j]
                artists[str(file1['artist_id'][j])]['genre'] = 'Unknown'
                artists[str(file1['artist_id'][j])]['start'] = 0000

            if i == 2:
                musics.append({'art_name' : file1['artist_names'][j],
                               'art_id' : file1['artists_id'][j],
                               'danceability' : file1['danceability'][j],
                               'energy' : file1['energy'][j],
                               'valence' : file1['valence'][j],
                               'tempo' : file1['tempo'][j],
                               'loudness' : file1['loudness'][j],
                               'mode' : file1['mode'][j],
                               'key' : file1['key'][j],
                               'acousticness' : file1['acousticness'][j],
                               'instrumentalness' : file1['instrumentalness'][j],
                               'liveness' : file1['liveness'][j],
                               'speechiness' : file1['speechiness'][j],
                               'duration_ms' : file1['duration_ms'][j],
                               'popularity' : file1['popularity'][j],
                               'year' : file1['year'][j],
                               'release_date' : file1['release_date'][j],
                               'song_title' : file1['song_title (censored)'][j],
                               })
                
            if i == 3:
                if str(file1['influencer_id'][j]) not in artists.keys():
                    print('new')
                    artists[str(file1['influencer_id'][j])]['name'] = file1['influencer_name'][j]
                    artists[str(file1['influencer_id'][j])]['fl_id'] = []
                artists[str(file1['influencer_id'][j])]['genre'] = file1['influencer_main_genre'][j]
                artists[str(file1['influencer_id'][j])]['start'] = file1['influencer_active_start'][j]
                artists[str(file1['influencer_id'][j])]['fl_id'].append(file1['follower_id'][j])
                if str(file1['follower_id'][j]) not in artists.keys():
                    artists[str(file1['follower_id'][j])] = {}
                    artists[str(file1['follower_id'][j])]['fl_id'] = []
                    artists[str(file1['follower_id'][j])]['name'] = ''
                    artists[str(file1['follower_id'][j])]['danceability'] = -10086
                    artists[str(file1['follower_id'][j])]['energy'] = -10086
                    artists[str(file1['follower_id'][j])]['valence'] = -10086
                    artists[str(file1['follower_id'][j])]['tempo'] = -10086
                    artists[str(file1['follower_id'][j])]['loudness'] = -10086
                    artists[str(file1['follower_id'][j])]['mode'] = -10086
                    artists[str(file1['follower_id'][j])]['key'] = -10086
                    artists[str(file1['follower_id'][j])]['acousticness'] = -10086
                    artists[str(file1['follower_id'][j])]['instrumentalness'] = -10086
                    artists[str(file1['follower_id'][j])]['liveness'] = -10086
                    artists[str(file1['follower_id'][j])]['speechiness'] = -10086
                    artists[str(file1['follower_id'][j])]['duration_ms'] = -10086
                    artists[str(file1['follower_id'][j])]['popularity'] = -10086
                    artists[str(file1['follower_id'][j])]['count'] = -10086
                artists[str(file1['follower_id'][j])]['genre'] = file1['follower_main_genre'][j]
                artists[str(file1['follower_id'][j])]['start'] = file1['follower_active_start'][j]

                follower_net[str(file1['follower_id'][j])] = {}
                follower_net[str(file1['follower_id'][j])]['name'] = []
                follower_net[str(file1['follower_id'][j])]['name'].append(file1['follower_name'][j])
                follower_net[str(file1['follower_id'][j])]['genre'] = file1['follower_main_genre'][j]
                follower_net[str(file1['follower_id'][j])]['start'] = file1['follower_active_start'][j]
                follower_net[str(file1['follower_id'][j])]['inf_id'] = file1['influencer_id'][j]
                follower_net[str(file1['follower_id'][j])]['inf_name'] = file1['influencer_name'][j]
                follower_net[str(file1['follower_id'][j])]['inf_genre'] = file1['influencer_main_genre'][j]
                follower_net[str(file1['follower_id'][j])]['inf_start'] = file1['influencer_active_start'][j]

    return artists, follower_net, musics


def make_subnet(start_id, net_file = r'F:\数模\比赛\influence net.xlsx'):
    start_id = int(start_id)
    file1 = pd.read_excel(net_file)
    ids = file1['ID']
    fl_idso = file1['FOLLOWER_IDS']
    fl_ids = []
    pass_list = ['[', ' ']
    for i in range(len(fl_idso)):
        fl_ids.append([])
        ll = ''
        for j in fl_idso[i]:
            if j in pass_list:
                continue
            elif j == ',' or j == ']':
                if ll != '':
                    fl_ids[i].append(ll)
                ll = ''
            else:
                ll += j

    print(ids[1392])
    print(start_id)
    st_infer = []
    st_fler = {}
    st_flids = []
    sec_infer = {}
    in_net = [str(start_id)]

    #从头开始往下找
    for i in range(len(ids)):
        if str(start_id) in fl_ids[i]:
            st_infer.append(str(ids[i]))
            in_net.append(str(ids[i]))
        if ids[i] == start_id:
            #print('found')
            for j in fl_ids[i]:
                st_fler[j] = {}
                sec_infer[j] = []
                if j not in in_net:
                    st_flids.append(str(j))
                    in_net.append(str(j))

    #向下扩展
    for i in st_flids:
        for j in range(len(ids)):
            if i == ids[j]:
                for k in fl_ids[j]:
                    st_fler[i][k] = {}
                    if k not in in_net:
                        st_flids.append(str(k))
                        in_net.append(str(k))
                    else:
                        for l in range(len(in_net)):
                            if in_net[l] == k:
                                st_fler[i][k] = {'back_to' : l}
                break

    #二层向上扩展
    for i in list(st_fler.keys()):
        if str(i) == str(start_id):
            continue
        for j in range(len(ids)):
            if str(j) == str(start_id):
                continue
            if i in fl_ids[j]:
                sec_infer[i].append(str(ids[j]))
                if i not in in_net:
                    in_net.append(str(i))

    return st_fler, st_infer, sec_infer, in_net


def get_genre(ids, rec_file = r'F:\数模\比赛\influence net.xlsx'):
    file1 = pd.read_excel(rec_file)
    art_mes = []
    genres = []
    for i in ids:
        #print(i)
        for j in range(len(file1['ID'])):
            if str(file1['ID'][j]) == str(i):
                art_mes.append({str(j) : [file1['NAME'][j], file1['GENRE'][j]]})
                break
    return art_mes


def dfs_fr(tr, col, lis, works):
    for i in tr.keys():
        works.write(col, lis, str(i))
        lis += 1 
        dfs_fr(tr[i], col+1, lis-1, works)
    return col-1, lis+1


def make_pic(fas, sons, nod, same_lv, file_path = r'F:\数模\比赛\2021_ICM_Problem_D_Data\2021_ICM_Problem_D_Data\influence_data.csv'):
    file1 = pd.read_csv(file_path)
    G=nx.DiGraph()
    ods = 1
    nods = [int(nod)]
    nods_name = []
    years = []
    genres = []
    vec = []
    vec_num = []
    vec_1 = []
    #加自身
    for j in range(len(file1['influencer_id'])):
        if str(nod) == str(file1['influencer_id'][j]):
            years.append(int(file1['influencer_active_start'][j]))
            genres.append(str(file1['influencer_main_genre'][j]))
            nods_name.append(str(file1['influencer_name'][j]))
            break

    #加父节点
    for i in fas:
        for j in range(len(file1['influencer_id'])):
            if str(i) == str(file1['influencer_id'][j]):
                years.append(int(file1['influencer_active_start'][j]))
                genres.append(str(file1['influencer_main_genre'][j]))
                nods.append(int(j))
                nods_name.append(str(file1['influencer_name'][j]))
                vec.append((nods_name[-1], nods_name[0]))
                if ((int(j), int(nod))) not in vec_num:
                    vec_num.append((int(j), int(nod)))
                break

    #加其他节点
    for i in list(same_lv.keys()):
        for j in range(len(file1['influencer_id'])):
            if str(i) == str(file1['follower_id'][j]):
                if int(i) not in nods:
                    nods.append(int(i))
                    years.append(int(file1['follower_active_start'][j]))
                    genres.append(str(file1['influencer_main_genre'][j]))
                    nods_name.append(str(file1['influencer_name'][j]))
                if ((int(nod), int(i))) not in vec_num:
                    vec_num.append((int(nod), int(i)))
                for l in range(len(nods)):
                    if nods[l] == int(i):
                        vec.append((nods_name[0], nods_name[l]))
                        break

        for j in same_lv[i]:
            for k in range(len(file1['influencer_id'])):
                if str(j) == str(file1['influencer_id'][k]):
                    if int(j) not in nods:
                        nods.append(int(j))
                        nods_name.append(str(file1['influencer_name'][k]))
                        years.append(int(file1['influencer_active_start'][k]))
                        genres.append(str(file1['influencer_main_genre'][k]))
                    if ((int(j), int(i))) not in vec_num:
                        vec_num.append((int(j), int(i)))
                    for l in range(len(nods)):
                        if nods[l] == int(i):
                            for m in range(len(nods)):
                                if nods[m] == int(j):
                                    vec.append((nods_name[m], nods_name[l]))
                                    break
                            break

    """G=nx.DiGraph()
    G.add_nodes_from(nods_name)
    G.add_edges_from(vec)
    nx.draw_networkx(G)"""

    G1=nx.DiGraph()
    #G1.add_nodes_from(nods)
    G1.add_edges_from(vec_num)
    nx.draw_networkx(G1)
    plt.show()

    return G1, nods, vec, vec_num, years, genres, nods_name, G1.number_of_edges()


def cal_param(nods, vec, years, genres, file_path = r'F:\数模\比赛\2021_ICM_Problem_D_Data\2021_ICM_Problem_D_Data\influence_data.csv'):
    file1 = pd.read_csv(file_path)
    new_vec = []
    old = 0
    new = 0
    for i in range(len(nods)):
        for j in range(i+1, len(nods)):
            new += 2
            if genres[i] == genres[j]:
                     new += 1

    for i in vec:
        i1 = -1
        i2 = -1
        for j in range(len(nods)):
            if int(nods[j]) == int(i[0]):
                i1 = j
            if int(nods[j]) == int(i[1]):
                i2 = j
            if i1 > -1 and i2 > -1:
                break
        if genres[i1] == genres[i2]:
            old += 1

    return (old + len(vec)) / new


def inf_data_get(artists, file_path = r'F:\数模\比赛\2021_ICM_Problem_D_Data\2021_ICM_Problem_D_Data\influence_data.csv'):
    file1 = pd.read_csv(file_path)
    fol_style_list = []
    inf_style_list = []
    lst = 0000
    ls_dic = {}
    for i in range(len(file1['follower_id'])):
        now = file1['follower_id'][i]
        now_i = file1['influencer_id'][i]
        if lst == now:
            fol_style_list.append(ls_dic)
            cnt = 1
            f_find = False
        else:
            cnt = 0
            f_find = True
        for j in list(artists.keys()):
            if f_find and str(now) == str(j):
                cnt += 1
                st_list = {}
                for k in list(artists[j].keys()):
                    st_list[str(k)] = artists[j][k]
                fol_style_list.append(st_list)
                ls_dic = st_list
            if str(now_i) == str(j):
                cnt += 1
                st_list = {}
                for k in list(artists[j].keys()):
                    st_list[str(k)] = artists[j][k]
                inf_style_list.append(st_list)
            if cnt == 2:
                break
    return inf_style_list,fol_style_list


if __name__ == "__main__":
    #读取数据
    #artists, inf_net, musics = load_data()
    year_list = [i for i in range(1920, 2021, 10)]
    key_list = ['name', 'fl_id', 'genre', 'start', 'danceability', 'energy', 'valence', 'tempo', 'loudness', 'mode', 'key', 'acousticness', 'instrumentalness', 'liveness', 'speechiness', 'duration_ms', 'popularity', 'count']
    cols1 = ['ID', 'NAME', 'FOLLOWER_IDS', 'GENRE', 'START', 'FOLLOWER_NUM', 'danceability', 'energy', 'valence', 'tempo', 'loudness', 'mode', 'key', 'acousticness', 'instrumentalness', 'liveness', 'speechiness', 'duration_ms', 'popularity', 'count']
    genres = {'Classical':{}, 'Pop/Rock':{}, 'Avant-Garde':{}, 'Easy Listening':{}, 'Latin':{}, 'Jazz':{}, 'Country':{}, 'R&B;':{}, 'Stage & Screen':{}, 'Vocal':{}}
    rootp = '180023'#'130028'#'418981'#'19241'#'180023'#'158888'
    rootp2 = '104741'
    genres_list = ['Classical', 'Pop/Rock', 'Avant-Garde', 'Easy Listening', 'Latin', 'Jazz', 'Country', 'R&B;', 'Stage & Screen', 'Vocal', 'Unknown', 'International', 'Reggae', 'Blues', 'Folk', 'Religious', 'Electronic', 'New Age', 'Comedy/Spoken', "Children's"]
    cont_ave = [31,
                955,
                5,
                16,
                92,
                144,
                157,
                210,
                21,
                117,
                43,
                27,
                28,
                24,
                27,
                13,
                25,
                5,
                7,
                1]
    for i in range(len(cont_ave)):
        print(genres_list[i], ':', cont_ave[i])
    #流派按年份总结
    """genre_year = {}
    for i in genres_list:
        genre_year[i] = {}
    song_genre = []
    for i in musics:
        for j in list(artists.keys()):
            st1= '[' + str(j)
            if st1 in str(i['art_id']):
                i['genre'] = artists[j]['genre']
                song_genre.append(i)
                break
    print(len(song_genre))
    ban_list = ['art_name', 'genre', 'art_id', 'year', 'release_date', 'song_title']
    for i in song_genre:
        if str(i['year']) not in list(genre_year[i['genre']].keys()):
            genre_year[i['genre']][str(i['year'])] = [1, i]
        else:
            for j in list(genre_year[i['genre']][str(i['year'])][1].keys()):
                if str(j) not in ban_list:
                    genre_year[i['genre']][str(i['year'])][1][j] += i[j]
            genre_year[i['genre']][str(i['year'])][0] += 1

    workbook = xlsxwriter.Workbook('genre_years.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 1
    i = song_genre[0]
    for j in list(genre_year[i['genre']][str(i['year'])][1].keys()):
        if str(j) not in ban_list:
            worksheet.write(0, lis, str(j))
            lis += 1
    col = 1
    for i in list(genre_year.keys()):
        lis = 0
        worksheet.write(col, lis, str(i))
        col += 1
        for j in genre_year[i].keys():
            lis = 0
            worksheet.write(col, lis,int(j))
            lis += 1
            for k in list(genre_year[i][j][1]):
                if str(k) not in ban_list:
                    if genre_year[i][j][0] != 1:
                        genre_year[i][j][1][k] /= genre_year[i][j][0]
                    worksheet.write(col, lis, genre_year[i][j][1][k])
                    lis += 1
            col += 1
    workbook.close()"""

    #搭建对应关系
    """stylei, stylef = inf_data_get(artists)
    workbook = xlsxwriter.Workbook('styles.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in key_list[4:-3]:
        worksheet.write(col, lis, 'i_'+i)
        lis += 1
    for i in key_list[4:-3]:
        worksheet.write(col, lis, 'f_'+i)
        lis += 1
    col = 1
    lis = 0
    for j in range(len(stylei)):
        for i in key_list[4:-3]:
            worksheet.write(col, lis, stylei[j][i])
            lis += 1
        col += 1
        lis = 0
    
    col = 1
    lis = len(key_list[4:-3])
    for j in range(len(stylef)):
        for i in key_list[4:-3]:
            worksheet.write(col, lis, stylef[j][i])
            lis += 1
        col += 1
        lis = len(key_list[4:-3])
    workbook.close()"""

    #第二问-2
    """counts_list = [3116, 97374, 481, 1680, 9418, 14733, 16014, 21387, 2104, 11941, 4401, 2773, 2859, 2495, 2790, 1365, 2505, 531, 664, 102]
    total = 0

    for i in counts_list:
        print(i /102)
        total += i
    print(total / 102)

    genre_inner = []
    cont_ave = [31,
                955,
                5,
                16,
                92,
                144,
                157,
                210,
                21,
                117,
                43,
                27,
                28,
                24,
                27,
                13,
                25,
                5,
                7,
                1]
    art_list = []
    for i in range(len(cont_ave)):
        art1 = []
        gen_in = []
        for j in range(cont_ave[i]):
            for k in list(artists.keys()):
                if str(artists[k]['genre']) == str(genres_list[i]) and str(k) not in art1:
                    art1.append(str(k))
                    g_i = []
                    for l in key_list[4:-3]:
                        g_i.append(artists[k][l])
                    gen_in.append(g_i)
                    break
        art_list.append(art1)
        genre_inner.append(gen_in)
    print(art_list)

    #随机取样计算
    for i in gen_in:

        #for j in random
        pass

    workbook = xlsxwriter.Workbook('arts.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in genres_list:
        worksheet.write(col, lis, i)
        col += 1
    col = 0
    lis = 1
    for i in art_list:
        for j in i:
            worksheet.write(col, lis, j)
            lis += 1
        col += 1
        lis = 1
    workbook.close()"""

    #画子网1
    """sons, fas, sec_fas, players = make_subnet(rootp)
    fas_mes = get_genre(fas)
    print(fas_mes)"""
    """jaz = 0
    pop = 0
    ele = 0
    reg = 0
    cla = 0
    sns = 0
    itn = 0
    oth = 0
    for i in fas_mes:
        if i[list(i.keys())[0]][1] == 'Jazz':
            jaz += 1
        elif i[list(i.keys())[0]][1] == 'Pop/Rock':
            pop += 1
        elif i[list(i.keys())[0]][1] == 'Electronic':
            ele += 1
        elif i[list(i.keys())[0]][1] == 'Reggae':
            reg += 1
        elif i[list(i.keys())[0]][1] == 'Classical':
            cla += 1
        elif i[list(i.keys())[0]][1] == 'Stage & Screen':
            sns += 1
        elif i[list(i.keys())[0]][1] == 'International':
            itn += 1
        else:
            oth += 1
    print(jaz, pop, ele, reg, cla, sns, itn, oth)"""

    """G, nods, vec, vec_num, years, genres, nods_name = make_pic(fas, sons, rootp, sec_fas)
    
    vec_n = []
    for i in vec_num:
        if i not in vec_n:
            vec_n.append(i)
    print(len(nods), len(vec_num))
    print(cal_param(nods, vec_num, years, genres))"""
    
    #画子网2
    """sons2, fas2, sec_fas2, players2 = make_subnet(rootp2)
    fas_mes2 = get_genre(fas2)
    print(fas_mes2)

    G, nods2, vec2, vec_num2, years2, genres2, nods_name2, num_of_vec = make_pic(fas2, sons2, rootp2, sec_fas2)
    
    vec_n2 = []
    for i in vec_num2:
        if i not in vec_n2:
            vec_n2.append(i)
    print(num_of_vec)
    print(len(nods2), len(vec_num2))
    print(cal_param(nods2, vec_num2, years2, genres2))"""

    #两张图合并
    """nods_t = nods
    years_t = years
    genres_t = genres
    for i in range(len(nods2)):
        if nods2[i] not in nods_t:
            nods_t.append(nods2[i])
            years_t.append(years2[i])
            genres_t.append(genres2[i])
    vec_t = vec_num
    for i in vec_num2:
        if i not in vec_t:
            vec_t.append(i)
    print(cal_param(nods_t, vec_t, years_t, genres_t))"""

    """print(sons.keys())
    workbook = xlsxwriter.Workbook('net2.xlsx')
    worksheet = workbook.add_worksheet()
    print(fas,'\n')
    col = 0
    lis = 0
    for i in fas:
        worksheet.write(col, lis, str(i))
        lis += 1
    col = 1
    lis = 0
    worksheet.write(col, lis, rootp)
    col = 2
    lis = 0
    dfs_fr(sons, col, lis, worksheet)
    col = 1
    lis = 2
    for i in sec_fas.keys():
        for j in sec_fas[i]:
            worksheet.write(col, lis, str(j))
            lis += 1
        lis += 1
        #print(str(i)+'`s fas:',sec_fas[i])

    workbook.close()"""

    #按流派总结
    """for i in genres.keys():
        for j in key_list[4:-3]:
            genres[i][j] = 0
        genres[i]['count'] = 0

    genre_by_years = {}
    for i in genres_list:
        genre_by_years[i] = {}
        for j in year_list:
            genre_by_years[i][j] = {}
            for k in key_list[4:-3]:
                genre_by_years[i][j][k] = -10086
    for i in artists.keys():
        if artists[i]['count'] == -10086:
            print(i, 'is lack of data')
            continue
        if 'genre' not in list(artists[i].keys()):
            print(artists[i].keys())
            continue


    for i in artists.keys():
        if artists[i]['count'] == -10086:
            print(i, 'is lack of data')
            continue
        if 'genre' not in list(artists[i].keys()):
            print(artists[i].keys())
            continue
        if artists[i]['genre'] not in genres:
            genres[artists[i]['genre']] = {}
            for j in key_list[4:-3]:
                genres[artists[i]['genre']][j] = float(artists[i][j]) * int(artists[i]['count'])
            genres[artists[i]['genre']]['count'] = int(artists[i]['count'])
        else:
            for j in key_list[4:-3]:
                genres[artists[i]['genre']][j] += float(artists[i][j]) * int(artists[i]['count'])
            genres[artists[i]['genre']]['count'] += int(artists[i]['count'])

    for i in genres.keys():
        if genres[i]['count'] == 0:
            print(i, 'is NONE')
            continue
        for j in key_list[4:-3]:
            genres[i][j] /= genres[i]['count']"""

    """workbook = xlsxwriter.Workbook('genres.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in genres['R&B;'].keys():
        worksheet.write(col, lis, str(i))
        lis += 1
    col = 1
    lis = 0
    for i in genres.keys():
        worksheet.write(col, lis, str(i))
        col += 1
    col = 0

    for i in genres.keys():
        col += 1
        lis = 1
        for j in key_list[4:-3]:
            worksheet.write(col, lis, str(genres[i][j]))
            lis += 1
    counts = [[], []]
    for i in genres.keys():
        print(i,':\t\t\t', genres[i]['count'])
        counts[0].append(genres[i]['count'])
        counts[1].append(i)
    workbook.close()
    print(counts[0])
    print(counts[1])"""

    #
    """workbook = xlsxwriter.Workbook('influence net.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in cols1[:6]:
        worksheet.write(col, lis, i)
        lis += 1
    for j in range(len(list(artists.keys()))):
        if artists[list(artists.keys())[j]]['fl_id'] == []:
            continue
        col += 1
        lis = 0
        worksheet.write(col, lis, list(artists.keys())[j])
        lis += 1
        same_genre = 0
        for i in ['name', 'fl_id', 'genre', 'start']:
            worksheet.write(col, lis, str(artists[list(artists.keys())[j]][i]))
            lis += 1
        worksheet.write(col, lis, len(artists[list(artists.keys())[j]]['fl_id']))

    workbook.close()"""


    """workbook = xlsxwriter.Workbook('influence net complex.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in cols1:
        worksheet.write(col, lis, i)
        lis += 1
    for j in range(len(list(artists.keys()))):
        if artists[list(artists.keys())[j]]['fl_id'] == []:
            continue
        col += 1
        lis = 0
        worksheet.write(col, lis, list(artists.keys())[j])
        lis += 1
        same_genre = 0
        for i in key_list[:4]:
            if i == 'fl_id':
                fls = []
                for k in artists[list(artists.keys())[j]][i]:
                    fls.append(str(k))
                fl = ', '.join(fls)
                worksheet.write(col, lis, fl)
                lis += 1
            else:
                worksheet.write(col, lis, str(artists[list(artists.keys())[j]][i]))
                lis += 1
        worksheet.write(col, lis, len(artists[list(artists.keys())[j]]['fl_id']))
        lis += 1
        for i in key_list[4:]:
            worksheet.write(col, lis, str(artists[list(artists.keys())[j]][i]))
            lis += 1

    workbook.close()"""


    """
    
    workbook = xlsxwriter.Workbook('.xlsx')
    worksheet = workbook.add_worksheet()
    col = 0
    lis = 0
    for i in cols1:
        worksheet.write(col, lis, i)
        lis += 1
    for j in range(len(list(artists.keys()))):
        if artists[list(artists.keys())[j]]['fl_id'] == []:
            continue
        col += 1
        lis = 0
        worksheet.write(col, lis, list(artists.keys())[j])
        lis += 1
        same_genre = 0
        for i in key_list[:4]:
            if i == 'fl_id':
                fls = []
                for k in artists[list(artists.keys())[j]][i]:
                    fls.append(str(k))
                fl = ', '.join(fls)
                worksheet.write(col, lis, fl)
                lis += 1
            else:
                worksheet.write(col, lis, str(artists[list(artists.keys())[j]][i]))
                lis += 1
        worksheet.write(col, lis, len(artists[list(artists.keys())[j]]['fl_id']))
        lis += 1
        for i in key_list[4:]:
            worksheet.write(col, lis, str(artists[list(artists.keys())[j]][i]))
            lis += 1

    workbook.close()"""

