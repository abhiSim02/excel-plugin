package com.poc.excelplugin.repository;

import com.poc.excelplugin.entity.UserFileHash;
import org.springframework.data.jpa.repository.JpaRepository;
import java.util.Optional;

public interface UserHashRepository extends JpaRepository<UserFileHash, Long> {
    Optional<UserFileHash> findByUserIdAndEntityName(String userId, String entityName);
}