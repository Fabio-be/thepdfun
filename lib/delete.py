import subprocess
import os

def delete_old_files(directory: str , minutes: int = 30) -> dict:
    """
    Supprime les fichiers plus vieux que X minutes dans un dossier
    
    Args:
        directory: Dossier à nettoyer
        minutes: Âge minimum en minutes (30 par défaut)
    
    Returns:
        dict: Résultat de la suppression
    """
    # Vérifier si le dossier existe
    if not os.path.exists(directory):
        return {"success": False, "error": f"Le dossier {directory} n'existe pas", "files_deleted": 0}
    
    # Compter les fichiers à supprimer
    count_cmd = ['find', directory, '-type', 'f', '-mmin', f'+{minutes}']
    count_result = subprocess.run(count_cmd, capture_output=True, text=True)
    files_to_delete = [f for f in count_result.stdout.strip().split('\n') if f]
    
    if not files_to_delete:
        return {"success": True, "message": f"Aucun fichier vieux de +{minutes} minutes", "files_deleted": 0}
    
    # Supprimer les fichiers
    delete_cmd = ['find', directory, '-type', 'f', '-mmin', f'+{minutes}', '-delete']
    result = subprocess.run(delete_cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        return {
            "success": True,
            "message": f"{len(files_to_delete)} fichier(s) supprimés",
            "files_deleted": len(files_to_delete),
            "directory": directory,
            "age_limit_minutes": minutes
        }
    else:
        return {
            "success": False,
            "error": result.stderr,
            "files_deleted": 0
        }

