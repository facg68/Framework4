		<%
			dim core_conexion, core_tt, core_sqlCommand, core_usuIndice, core_WallPaper

			set core_conexion = Server.CreateObject("ADODB.Connection")
			core_conexion.open Application("Conn")	

            function limpiar(cadena)
                limpiar = Replace(cadena,"'","´")    
            end function    

			function UsuarioExiste(usuario)
				dim core_sqlCommand

				core_sqlCommand = "SELECT COUNT(*) AS Existe " & _
							   "FROM seg_Usuarios " & _
							  "WHERE (usuCodigo = '" & Usuario & "') " & _
								"AND (usuEstado = 1);"

				set core_tt = core_conexion.execute(core_sqlCommand)
					UsuarioExiste = core_tt("Existe")
				core_tt.close: set core_tt = nothing
			end function

			function NombreUsuario(fxUsuario)	
				set core_tt = core_conexion.execute("SELECT usuNombre FROM seg_Usuarios WHERE usuCodigo = '" & fxUsuario & "';")        
					if not (core_tt.bof or core_tt.eof) then
						NombreUsuario = core_tt("usuNombre")
					end if
				core_tt.close: set core_tt = nothing
			end function				

			function ResetearPassword(usuario)
				set core_tt = core_conexion.execute("SELECT usuReset FROM seg_Usuarios WHERE usuCodigo = '" & Usuario & "';")
					ResetearPassword = core_tt("usuReset")
				core_tt.close: set core_tt = nothing
			end function

			function CadenaRandom()
				dim CharacterSetArray, TempCopy, i, j, Count, Chars, Index,Temp

				Randomize()

				CharacterSetArray = Array(_
					Array(7, "abcdefghijklmnopqrstuvwxyz"), _
					Array(1, "0123456789") _
				)

				for i = 0 to UBound(CharacterSetArray)
					Count = CharacterSetArray(i)(0)
					Chars = CharacterSetArray(i)(1)

					for j = 1 to Count
						Index = Int(Rnd() * Len(Chars)) + 1
						Temp = Temp & Mid(Chars, Index, 1)
					next
				next

				do until Len(Temp) = 0
					Index = Int(Rnd() * Len(Temp)) + 1
					TempCopy = TempCopy & Mid(Temp, Index, 1)
					Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)
				loop

				CadenaRandom = TempCopy
			end function

			sub CrearMenu(fxUsuario)
				core_usuIndice = 0
				core_conexion.execute("DELETE FROM seg_WebMenu WHERE Usuario = '" & fxUsuario & "';")

				InsertarComando fxUsuario, "<li><a class=|mb| href=|/core/gohome.asp| class=|nav-active|>&#9733;</a></li>", 1
				DibujarMenuBar fxUsuario
			end sub

			sub ActualizarSnippetsUsuario(fxUsuario)
				SQLString = "INSERT INTO seg_Usuarios_Snippets(codUsuario, codSistema, codProceso, Snippet, snippetActivo) " & _
							"SELECT '" & fxUsuario & "' AS Usuario, proSistema, proCodigo, Snippet, Estatus " & _
							"FROM (SELECT pr.proSistema, pr.proCodigo, pr.Snippet, 1 AS Estatus " & _
										"FROM dbo.seg_PermisosUsuarios AS pu " & _
										"INNER JOIN dbo.seg_Procesos AS pr " & _
										"ON pu.Sistema = pr.proSistema " & _
										"AND pu.Proceso = pr.proCodigo " & _
										"WHERE (pr.proMenuItem = 1) " & _
										"AND (pu.Usuario = '" & fxUsuario & "') " & _
										"AND (pr.Snippet IS NOT NULL)" & _
							     ") AS SnippetsValidos " & _
							"WHERE (Snippet NOT IN (SELECT snippet FROM dbo.seg_Usuarios_Snippets WHERE (codUsuario = '" & fxUsuario & "')))"

				core_conexion.execute(SQLString)
			end sub			
			
			sub DibujarMenuBar(fxUsuario)
				'
				' Dibujamos los Items "Root" que el usuario
				' puede ver (Menú Raíz) 
				'
				dim t, vinculo, cmdString

				cmdString  = "exec dbo.seg_Web_MenuBar '" & fxUsuario & "', 1;"    
				set t = core_conexion.execute(cmdString)

				if not (t.bof or t.eof) then
					do
						if t("pro_Items") > 0 then
							InsertarComando fxUsuario, "<li><a class=|mb| href=|#| class=|nav-active|>" & t("sysNombre") & "</a>", 0
								InsertarComando fxUsuario, "<ul>", 0
									DibujarNivel fxUsuario, t("proSistema"), t("sysCodigo")
								InsertarComando fxUsuario, "</ul>", 0
							InsertarComando fxUsuario, "</li>", 0
						else
							InsertarComando fxUsuario, "<li><a class=|mb| href=|#| class=|nav-active|>" & t("sysNombre") & "</a></li>", 0
						end if

						t.movenext
					loop until t.eof
				end if

				t.close: set t = nothing
			end sub

			sub DibujarNivel(fxUsuario, Sistema, Codigo)
				'
				' Dibujamos la(s) opción(es) que el usuario tiene
				' asignados en el menú principal en análisis
				'
				dim sm, vinculo, cmdString

				cmdString = "exec seg_Web_MenuBar_Niveles '" & fxUsuario & "' ,'" & Codigo & "'"
				set sm = core_conexion.execute(cmdString)

				if not (sm.bof or sm.eof) then
					do
						if MenuItems(fxUsuario, sm("Proceso")) > 0 then
							InsertarComando fxUsuario, "<li><a class=|mb| href=|#| class=|nav-active|>" & sm("proNombre") & "</a>" , 0
								InsertarComando fxUsuario, "<ul>" , 0      
									DibujarNivel fxUsuario, Sistema, sm("Proceso")
								InsertarComando fxUsuario, "</ul>", 0
							InsertarComando fxUsuario, "</li>" , 0         
						else
							if (left(uCase(sm("proAction")), 4) = "Hcore_ttP") or (left(uCase(sm("proAction")), 3) = "WWW") then
								vinculo = sm("proAction")
								if len(trim(sm("proActionParam"))) > 0 then
									vinculo = vinculo & sm("proActionParam")
								end if								
							else
								vinculo = "/forma/desktop/apps/" & Sistema & "/" & sm("proAction") & ".asp"
								if len(trim(sm("proActionParam"))) > 0 then
									vinculo = vinculo & sm("proActionParam")
								end if								
							end if

							if sm("Externo") = 0 then
								InsertarComando fxUsuario, "<li><a class=|mb| href=|" & vinculo & "| class=|nav-active|>" & sm("proNombre") & "</a></li>", 0
							else
								InsertarComando fxUsuario, "<li><a class=|mb| href=|" & vinculo & "| class=|nav-active| target=|_blank| rel=|noopener noreferrer|>" & sm("proNombre") & "</a></li>", 0
							end if							
						end if

						sm.movenext
					loop until sm.eof
				end if

				sm.close: set sm = nothing   
			end sub			

			function MenuItems(fxUsuario, Codigo)
				core_sqlCommand = "exec seg_Web_MenuBar_Items '" & fxUsuario & "', '" & Codigo & "'"

				set core_tt = core_conexion.execute(core_sqlCommand)
					if not (core_tt.bof or core_tt.eof) then
						MenuItems = core_tt("Items")
					else
						MenuItems = 0
					end if
				core_tt.close: set core_tt = nothing
			end function

			sub InsertarComando(fxUsuario, Comando, switch)
				dim usuMenu

				core_usuIndice = core_usuIndice + 1
				if len(switch) = 0 then switch = 0

				usuMenu = "INSERT INTO seg_WebMenu(Usuario, Indice, Comando, home) " & _
						  "VALUES ('" & fxUsuario & "', " & core_usuIndice & ", '" & Comando & "', " & switch & ");"

				core_conexion.execute(usuMenu)		
			end sub		

			Function ContarWallpapers()
				dim fs, imagen, strFilePath

				set fs = Server.CreateObject("Scripting.FileSystemObject")
				ContarWallpapers = 0

				do while (True)
					ContarWallpapers = ContarWallpapers + 1
					imagen = right("00000000" & cStr(ContarWallpapers), 8) & ".jpg"
					strFilePath = Server.MapPath("/") & "\core\imagenes\fondos\" & imagen 

					if not (fs.FileExists(strFilePath)) then 
						exit do
					end if
				loop

				ContarWallpapers = ContarWallpapers - 1

				set fs = nothing		
			end Function	

            function wallPaperUsuario()
                core_sqlCommand = "SELECT usuWallPaper FROM seg_Usuarios WHERE usuCodigo = '" & Request.Cookies("Usuario") & "';"
                core_WallPaper = 1

                set core_tt = core_conexion.execute(core_sqlCommand)

                if not (core_tt.bof or core_tt.eof) then
                    if len(trim(core_tt("usuWallPaper"))) > 0 then
                        wallPaperUsuario = cInt(left(core_tt("usuWallPaper"), 8))
                    end if
                end if

                core_tt.close: set core_tt = nothing
            end function   			

            function usuHomePage(usuario)
                set core_tt = core_conexion.execute("SELECT usuHomePage FROM seg_Usuarios WHERE usuCodigo = '" & usuario & "';")
                    usuHomePage = core_tt("usuHomePage")
                core_tt.close: set core_tt = nothing
            end function			

            sub ResetearHomePage(usuario)	
                core_conexion.execute("UPDATE seg_Usuarios SET usuHomePage = NULL WHERE usuCodigo = '" & usuario & "';")        
            end sub  										
		%>