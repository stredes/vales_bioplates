#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Gestión de solicitantes y usuarios de bodega."""

from __future__ import annotations

import json
import os
from typing import List


class UserManager:
    """Gestiona listas de solicitantes y usuarios de bodega."""
    
    def __init__(self, data_dir: str = ".") -> None:
        self.data_dir = data_dir
        self.users_file = os.path.join(data_dir, 'usuarios_bodega.json')
        self.solicitantes_file = os.path.join(data_dir, 'solicitantes.json')
        self._load()
    
    def _load(self) -> None:
        """Carga las listas desde archivos JSON."""
        # Cargar usuarios bodega
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'r', encoding='utf-8') as f:
                    self.usuarios_bodega = json.load(f)
            except Exception:
                self.usuarios_bodega = []
        else:
            self.usuarios_bodega = []
        
        # Cargar solicitantes
        if os.path.exists(self.solicitantes_file):
            try:
                with open(self.solicitantes_file, 'r', encoding='utf-8') as f:
                    self.solicitantes = json.load(f)
            except Exception:
                self.solicitantes = []
        else:
            self.solicitantes = []
    
    def _save_usuarios(self) -> None:
        """Guarda la lista de usuarios bodega."""
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            with open(self.users_file, 'w', encoding='utf-8') as f:
                json.dump(self.usuarios_bodega, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise Exception(f"Error al guardar usuarios: {e}")
    
    def _save_solicitantes(self) -> None:
        """Guarda la lista de solicitantes."""
        try:
            os.makedirs(self.data_dir, exist_ok=True)
            with open(self.solicitantes_file, 'w', encoding='utf-8') as f:
                json.dump(self.solicitantes, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise Exception(f"Error al guardar solicitantes: {e}")
    
    def get_usuarios_bodega(self) -> List[str]:
        """Retorna la lista de usuarios de bodega."""
        return self.usuarios_bodega.copy()
    
    def get_solicitantes(self) -> List[str]:
        """Retorna la lista de solicitantes."""
        return self.solicitantes.copy()
    
    def add_usuario_bodega(self, nombre: str) -> bool:
        """Agrega un usuario de bodega si no existe."""
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío")
        if nombre in self.usuarios_bodega:
            return False
        self.usuarios_bodega.append(nombre)
        self._save_usuarios()
        return True
    
    def add_solicitante(self, nombre: str) -> bool:
        """Agrega un solicitante si no existe."""
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío")
        if nombre in self.solicitantes:
            return False
        self.solicitantes.append(nombre)
        self._save_solicitantes()
        return True
    
    def remove_usuario_bodega(self, nombre: str) -> bool:
        """Elimina un usuario de bodega."""
        if nombre in self.usuarios_bodega:
            self.usuarios_bodega.remove(nombre)
            self._save_usuarios()
            return True
        return False
    
    def remove_solicitante(self, nombre: str) -> bool:
        """Elimina un solicitante."""
        if nombre in self.solicitantes:
            self.solicitantes.remove(nombre)
            self._save_solicitantes()
            return True
        return False
